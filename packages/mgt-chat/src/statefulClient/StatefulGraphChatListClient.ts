/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MessageThreadProps, ErrorBarProps, Message } from '@azure/communication-react';
import { ActiveAccountChanged, IGraph, Providers, log } from '@microsoft/mgt-element';
import { GraphError } from '@microsoft/microsoft-graph-client';
import {
  AadUserConversationMember,
  ChatMessage,
  ChatRenamedEventMessageDetail,
  ItemBody,
  MembersAddedEventMessageDetail,
  MembersDeletedEventMessageDetail,
  TeamsAppInstalledEventMessageDetail,
  TeamsAppRemovedEventMessageDetail
} from '@microsoft/microsoft-graph-types';
import { produce } from 'immer';
import { currentUserId } from '../utils/currentUser';
import { graph } from '../utils/graph';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationUserClient } from './GraphNotificationUserClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { ChatThreadCollection, loadChatThreads, loadChatThreadsByPage } from './graph.chat';
import { ChatMessageInfo, Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { error } from '@microsoft/mgt-element';
import { LastReadCache } from '../statefulClient/Caching/LastReadCache';

interface ODataType {
  '@odata.type': MessageEventType;
}
type MembersAddedEventDetail = ODataType &
  MembersAddedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.membersAddedEventMessageDetail';
  };
type MembersRemovedEventDetail = ODataType &
  MembersDeletedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.membersDeletedEventMessageDetail';
  };
type ChatRenamedEventDetail = ODataType &
  ChatRenamedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.chatRenamedEventMessageDetail';
  };
type TeamsAppInstalledEventDetail = ODataType &
  TeamsAppInstalledEventMessageDetail & {
    '@odata.type': '#microsoft.graph.teamsAppInstalledEventMessageDetail';
  };
type TeamsAppRemovedEventDetail = ODataType &
  TeamsAppRemovedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.teamsAppRemovedEventMessageDetail';
  };

type ChatMessageEvents =
  | MembersAddedEventDetail
  | MembersRemovedEventDetail
  | ChatRenamedEventDetail
  | TeamsAppInstalledEventDetail
  | TeamsAppRemovedEventDetail;

export type GraphChatThread = GraphChat & {
  isRead: boolean;
};

// defines the type of the state object returned from the StatefulGraphChatListClient
export type GraphChatListClient = Pick<MessageThreadProps, 'userId'> & {
  status:
    | 'initial'
    | 'creating server connections'
    | 'server connection lost'
    | 'server connection established'
    | 'subscribing to notifications'
    | 'loading chats'
    | 'no chats'
    | 'chats loaded'
    | 'chat message received'
    | 'chats read'
    | 'chat selected'
    | 'chat unselected';
  chatThreads: GraphChatThread[];
  chatMessage: ChatMessage | undefined;
  moreChatThreadsToLoad: boolean | undefined;
  internalSelectedChat: GraphChatThread | undefined;
  internalPrevSelectedChat: GraphChatThread | undefined;
  initialSelectedChatSet: boolean;
} & Pick<ErrorBarProps, 'activeErrorMessages'>;

interface StatefulClient<T> {
  /**
   * Get the current state of the client
   */
  getState(): T;
  /**
   * Register a callback to receive state updates
   *
   * @param handler Callback to receive state updates
   */
  onStateChange(handler: (state: T) => void): void;
  /**
   * Remove a callback from receiving state updates
   *
   * @param handler Callback to be unregistered
   */
  offStateChange(handler: (state: T) => void): void;

  chatThreadsPerPage: number;
  /**
   * Method for loading more chat threads
   */
  loadMoreChatThreads(): void;
  /**
   * Method for marking all chat threads as read
   */
  markAllChatThreadsAsRead(): void;
  /**
   * Method for caching last read time for all included chat threads
   */
  cacheLastReadTime(action: 'all' | 'selected'): void;
  /**
   * Method for setting the currently selected chat
   * @param chatThread - currently selected chat
   */
  setInternalSelectedChat(chatThread: GraphChatThread): void;
}

type MessageEventType =
  | '#microsoft.graph.membersAddedEventMessageDetail'
  | '#microsoft.graph.membersDeletedEventMessageDetail'
  | '#microsoft.graph.chatRenamedEventMessageDetail'
  | '#microsoft.graph.teamsAppInstalledEventMessageDetail'
  | '#microsoft.graph.teamsAppRemovedEventMessageDetail';

/**
 * Extended Message type with additional properties.
 */
export type GraphChatMessage = Message & {
  hasUnsupportedContent: boolean;
  rawChatUrl: string;
};

export interface ChatListEvent {
  type:
    | 'chatMessageReceived'
    | 'chatMessageDeleted'
    | 'chatMessageEdited'
    | 'chatRenamed'
    | 'memberAdded'
    | 'memberRemoved'
    | 'systemEvent'
    | 'teamsAppInstalled'
    | 'teamsAppRemoved';
  message: ChatMessage;
}

class StatefulGraphChatListClient implements StatefulClient<GraphChatListClient> {
  private readonly _notificationClient: GraphNotificationUserClient;
  private readonly _eventEmitter: ThreadEventEmitter;
  private readonly _cache: LastReadCache;
  private readonly _graph: IGraph;
  private _stateSubscribers: ((state: GraphChatListClient) => void)[] = [];
  private _initialSelectedChatId: string | undefined;
  constructor(chatThreadsPerPage: number, initialSelectedChatId: string | undefined) {
    this.userId = currentUserId();
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    this._cache = new LastReadCache();
    this._graph = graph('mgt-chat', GraphConfig.version);
    this.chatThreadsPerPage = chatThreadsPerPage;
    this._initialSelectedChatId = initialSelectedChatId;
    this._notificationClient = new GraphNotificationUserClient(this._eventEmitter, this._graph);

    void this.updateUserSubscription(this.userId);
  }

  /**
   * Provides the number of chat threads to display with each load more.
   */
  public chatThreadsPerPage: number;

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public tearDown() {
    this._notificationClient.tearDown();
  }

  /**
   * Load more chat threads if applicable.
   */
  public async loadMoreChatThreads() {
    const state = this.getState();
    const items: GraphChatThread[] = [];
    await this.loadAndAppendChatThreads('', items, state.chatThreads.length + this.chatThreadsPerPage);
  }

  private async loadAndAppendChatThreads(nextLink: string, items: GraphChatThread[], maxItems: number) {
    if (maxItems < 1) {
      error('maxItem is invalid: ' + maxItems);
      return;
    }

    const handler = async (latestChatThreads: ChatThreadCollection) => {
      const latestItems = (latestChatThreads.value as GraphChatThread[]).filter(chatThread => chatThread.id);
      const checkedItems = await this.checkWhetherToMarkAsRead(latestItems);
      items = items.concat(checkedItems);

      const handlerNextLink = latestChatThreads['@odata.nextLink'];

      if (items.length > maxItems) {
        // return exact page size
        this.handleChatThreads(items.slice(0, maxItems), 'more');
        return;
      }

      if (items.length < maxItems && handlerNextLink && handlerNextLink !== '') {
        await this.loadAndAppendChatThreads(handlerNextLink, items, maxItems);
        return;
      }

      this.handleChatThreads(items, handlerNextLink);
    };

    if (nextLink === '') {
      // max page count cannot exceed 50 per documentation
      const pageCount = maxItems > 50 ? 50 : maxItems;
      loadChatThreads(this._graph, pageCount).then(handler, err => error(err));
    } else {
      const filter = nextLink.split('?')[1];
      await loadChatThreadsByPage(this._graph, filter).then(handler, err => error(err));
    }
  }

  public setInternalSelectedChat = (chatThread: GraphChatThread): void => {
    const state = this.getState();

    if (state.internalSelectedChat) {
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'chat unselected';
        draft.internalPrevSelectedChat = state.internalSelectedChat;
        draft.internalSelectedChat = undefined;
      });
    }

    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'chat selected';
      draft.internalSelectedChat = {
        ...chatThread,
        isRead: true
      };
    });
  };

  public markAllChatThreadsAsRead = () => {
    // mark as read after chat thread is found in current state
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'chats read';
      draft.chatThreads = this._state.chatThreads.map((chatThread: GraphChatThread) => {
        if (chatThread.id && !chatThread.isRead) {
          return {
            ...chatThread,
            isRead: true
          };
        }
        return chatThread;
      });
    });

    this.cacheLastReadTime('all');
  };

  public cacheLastReadTime = (action: 'all' | 'selected') => {
    const state = this.getState();
    // cache last read time after all chat threads
    if (action === 'all') {
      state.chatThreads.forEach((chatThread: GraphChatThread) => {
        if (chatThread.id) {
          void this._cache.cacheLastReadTime(chatThread.id, new Date());
        }
      });
    }

    // cache last read time after chat thread is found in current state
    if (action === 'selected') {
      if (state.internalSelectedChat?.id) {
        log(`caching the last-read timestamp of now to chat ID '${state.internalSelectedChat.id}'...`);
        void this._cache.cacheLastReadTime(state.internalSelectedChat.id, new Date());
      }
    }
  };

  // check whether to mark the chat as read or not
  private readonly checkWhetherToMarkAsRead = async (c: GraphChatThread[]): Promise<GraphChatThread[]> => {
    const result = await Promise.all(
      c.map(async (chatThread: GraphChatThread) => {
        const lastReadData = await this._cache.loadLastReadTime(chatThread.id!);
        if (lastReadData) {
          const lastUpdatedDateTime = chatThread.lastUpdatedDateTime ? new Date(chatThread.lastUpdatedDateTime) : null;
          const lastMessagePreviewCreatedDateTime = chatThread.lastMessagePreview?.createdDateTime
            ? new Date(chatThread.lastMessagePreview?.createdDateTime)
            : null;
          const lastReadTime = new Date(lastReadData.lastReadTime);
          const isRead = !(
            (lastUpdatedDateTime && lastUpdatedDateTime > lastReadTime) ||
            (lastMessagePreviewCreatedDateTime && lastMessagePreviewCreatedDateTime > lastReadTime) ||
            !lastReadData.lastReadTime
          );
          return {
            ...chatThread,
            isRead
          };
        } else {
          return chatThread;
        }
      })
    );
    return result;
  };

  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatListClient) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public onStateChange(handler: (state: GraphChatListClient) => void): void {
    if (!this._stateSubscribers.includes(handler)) {
      this._stateSubscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving state updates
   *
   * @param {(state: GraphChatListClient) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public offStateChange(handler: (state: GraphChatListClient) => void): void {
    const index = this._stateSubscribers.indexOf(handler);
    if (index !== -1) {
      this._stateSubscribers = this._stateSubscribers.splice(index, 1);
    }
  }

  private readonly _initialState: GraphChatListClient = {
    status: 'initial',
    activeErrorMessages: [],
    userId: '',
    chatThreads: [],
    moreChatThreadsToLoad: undefined,
    chatMessage: undefined,
    internalSelectedChat: undefined,
    internalPrevSelectedChat: undefined,
    initialSelectedChatSet: false
  };

  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatListClient}
   * @memberof StatefulGraphChatListClient
   */
  private _state: GraphChatListClient = { ...this._initialState };

  /**
   * Calls each subscriber with the next state to be emitted
   *
   * @param recipe - a function which produces the next state to be emitted
   */
  private notifyStateChange(recipe: (draft: GraphChatListClient) => void) {
    this._state = produce(this._state, recipe);
    this._stateSubscribers.forEach(handler => handler(this._state));
  }

  /**
   * Handle ChatListEvent event types.
   */
  private notifyChatMessageEventChange(event: ChatListEvent) {
    this.notifyStateChange((draft: GraphChatListClient) => {
      // find the chat thread
      const chatThreadIndex = draft.chatThreads.findIndex(c => c.id === event.message.chatId);
      const chatThread: GraphChatThread | undefined = draft.chatThreads[chatThreadIndex];

      if (
        chatThread &&
        event.message.lastModifiedDateTime &&
        chatThread.lastMessagePreview?.createdDateTime &&
        event.message.lastModifiedDateTime < chatThread.lastMessagePreview.createdDateTime
      ) {
        log('message is older than last message preview, ignoring', event.message);
        return;
      }

      // func to bring the chat thread to the top of the list
      const bringToTop = (newThread?: GraphChatThread) => {
        draft.chatThreads.splice(chatThreadIndex, 1);
        if (newThread) {
          draft.chatThreads.unshift(newThread);
        } else if (chatThread) {
          draft.chatThreads.unshift(chatThread);
        }
      };

      // handle the events
      if (event.type === 'chatRenamed' && event.message?.eventDetail && chatThread) {
        // rename the chat thread in-place
        chatThread.topic =
          (event.message.eventDetail as ChatRenamedEventMessageDetail)?.chatDisplayName ?? chatThread.topic;
      } else if (event.type === 'memberAdded' && chatThread) {
        // inject a chat thread with only the id to force a reload; this is necessary because the
        // notification does not include the displayName of the user who was added
        const newThread = { id: chatThread.id } as GraphChatThread;
        bringToTop(newThread);
      } else if (event.type === 'memberRemoved' && event.message?.eventDetail && chatThread) {
        // update the user list to remove the user; while we could add a "User removed" message
        // the Teams client doesn't show a message and the Graph API does not show a message either
        const membersToRemove = (event.message.eventDetail as MembersDeletedEventMessageDetail).members;
        if (membersToRemove) {
          const membersToRemoveSet = new Set(membersToRemove.map(member => member.id));
          chatThread.members = chatThread.members?.filter(member => {
            const user = member as AadUserConversationMember;
            return !membersToRemoveSet.has(user.userId);
          });
        }
      } else if (event.type === 'teamsAppInstalled' && event.message?.eventDetail && chatThread) {
        // add a message about the app being added; Teams does "User added", but this provides
        // more clarity since the user list doesn't change
        const details = event.message.eventDetail as TeamsAppInstalledEventMessageDetail;
        chatThread.lastMessagePreview = {
          body: {
            content: `${details.teamsAppDisplayName || details.teamsAppId} app added`,
            contentType: 'text'
          } as ItemBody,
          lastModifiedDateTime: event.message.lastModifiedDateTime
        } as ChatMessageInfo;
        log('injected', chatThread.lastMessagePreview);
        bringToTop();
      } else if (event.type === 'teamsAppRemoved' && event.message?.eventDetail && chatThread) {
        // there is no behavior in teams for an app being removed
      } else if (event.type === 'chatMessageEdited' && chatThread) {
        // update the last message preview in-place
        chatThread.lastMessagePreview = event.message as ChatMessageInfo;
        chatThread.lastUpdatedDateTime = event.message.lastModifiedDateTime;
      } else if (event.type === 'chatMessageDeleted' && chatThread) {
        // update the last message preview in-place
        chatThread.lastMessagePreview = {
          lastModifiedDateTime: event.message.lastModifiedDateTime,
          isDeleted: true
        } as ChatMessageInfo;
      } else if (event.type === 'chatMessageReceived' && event.message && chatThread) {
        draft.status = 'chat message received';
        draft.chatMessage = event.message;
        // update the last message preview and bring to the top
        chatThread.lastMessagePreview = event.message as ChatMessageInfo;
        chatThread.lastUpdatedDateTime = event.message.lastModifiedDateTime;
        // this resets the chat thread read state for all chats including the active chat
        if (draft.internalSelectedChat && draft.internalSelectedChat.id === draft.chatMessage.chatId) {
          chatThread.isRead = true;
        } else {
          chatThread.isRead = false;
        }

        bringToTop();
      } else if (event.type === 'chatMessageReceived' && event.message?.chatId) {
        draft.status = 'chat message received';
        draft.chatMessage = event.message;
        // create a new chat thread at the top
        const newChatThread: GraphChatThread = {
          id: event.message.chatId,
          lastMessagePreview: event.message as ChatMessageInfo,
          lastUpdatedDateTime: event.message.lastModifiedDateTime,
          isRead: false
        };
        draft.chatThreads.unshift(newChatThread);
      } else {
        log(`received unrecognized event type '${event.type}' from the user subscription.`);
      }
    });
  }

  /*
   * Returns the type of events by checking the chat message.
   */
  private getSystemMessageType(message: ChatMessage) {
    // check if this is a SystemEvent
    if (message.messageType === 'systemEventMessage') {
      const eventDetail = message.eventDetail as ChatMessageEvents;
      switch (eventDetail['@odata.type']) {
        case '#microsoft.graph.membersAddedEventMessageDetail':
          return 'memberAdded';
        case '#microsoft.graph.membersDeletedEventMessageDetail':
          return 'memberRemoved';
        case '#microsoft.graph.chatRenamedEventMessageDetail':
          return 'chatRenamed';
        case '#microsoft.graph.teamsAppInstalledEventMessageDetail':
          return 'teamsAppInstalled';
        case '#microsoft.graph.teamsAppRemovedEventMessageDetail':
          return 'teamsAppRemoved';
        default:
          return eventDetail['@odata.type'];
      }
    }

    return 'chatMessageReceived';
  }

  /*
   * Event handler to be called when a new message is received by the notification service
   */
  private readonly onMessageReceived = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: this.getSystemMessageType(message) });
    }
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private readonly onMessageDeleted = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageDeleted' });
    }
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private readonly onMessageEdited = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageEdited' });
    }
  };

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatListClient}
   * @memberof StatefulGraphChatListClient
   */
  public getState(): GraphChatListClient {
    return this._state;
  }

  /*
   * Event handler to be called when we need to load more chat threads.
   */
  private readonly handleChatThreads = (chatThreads: GraphChatThread[], nextLink: string | undefined) => {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = chatThreads.length > 0 ? 'chats loaded' : 'no chats';
      draft.chatThreads = chatThreads;
      draft.moreChatThreadsToLoad = nextLink !== undefined && nextLink !== '';
      if (this._initialSelectedChatId) {
        draft.internalSelectedChat = chatThreads.filter(c => c.id === this._initialSelectedChatId)[0];
        this._initialSelectedChatId = undefined;
        draft.initialSelectedChatSet = true;
      }
    });
  };

  private readonly onActiveAccountChanged = (e: ActiveAccountChanged) => {
    if (!e.detail) {
      return;
    }
    const [newUserId] = e.detail.id.split('.');

    if (newUserId && this.userId !== newUserId) {
      void this.handleAccountChange(newUserId);
    }
  };

  private readonly handleAccountChange = async (userId: string) => {
    await this._notificationClient.unsubscribeFromUserNotifications(this.userId);

    this.clearCurrentUserMessages();

    this.userId = userId;
    // by updating the followed chat the notification client will reconnect to SignalR
    await this.updateUserSubscription(userId);
  };

  private clearCurrentUserMessages() {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'initial';
      draft.chatThreads = [];
    });
  }

  /**
   * Current User ID.
   */
  private _userId = '';

  /**
   * Returns the current User ID.
   */
  public get userId() {
    return this._userId;
  }

  /**
   * Sets the current User ID and updates the state value.
   */
  private set userId(userId: string) {
    if (this._userId === userId) {
      return;
    }
    this._userId = userId;
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.userId = userId;
    });
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatListClient
   */
  private async updateUserSubscription(userId: string) {
    if (userId === '') return;

    // reset state to initial
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'initial';
    });
    // Subscribe to notifications for messages
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'creating server connections';
    });
    try {
      await this._notificationClient.subscribeToUserNotifications(userId);
    } catch (e) {
      error('Failed to load chat data or subscribe to notications: ', e);
      if (e instanceof GraphError) {
        this.notifyStateChange((draft: GraphChatListClient) => {
          draft.status = 'no chats';
        });
      }
    }
  }

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', (message: ChatMessage) => void this.onMessageReceived(message));
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', (message: ChatMessage) => void this.onMessageEdited(message));
    this._eventEmitter.on('disconnected', () => {
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'server connection lost';
        draft.chatThreads = [];
      });
    });
    this._eventEmitter.on('connected', () => {
      void this.loadAndAppendChatThreads('', [], this.chatThreadsPerPage);
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'server connection established';
      });
    });
  }
}

export { StatefulGraphChatListClient };
