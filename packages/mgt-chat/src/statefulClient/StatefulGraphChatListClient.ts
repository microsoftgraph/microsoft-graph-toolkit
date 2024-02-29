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
import { ChatThreadCollection, loadChatThreads, loadChatThreadsByPage, loadChatWithPreview } from './graph.chat';
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
    | 'chat details loaded'
    | 'chats read'
    | 'chat selected'
    | 'chat unselected'
    | 'fatal error';
  chatThreads: GraphChatThread[];
  chatMessage: ChatMessage | undefined;
  moreChatThreadsToLoad: boolean | undefined;
  internalSelectedChat: GraphChatThread | undefined;
  internalPrevSelectedChat: GraphChatThread | undefined;
  fireOnSelected: boolean; // takes care of a case when we first init ChatList and sets the selected chat Id but onselected is not fired.
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
  /**
   * Provides the number of chat threads to display with each load more.
   */
  chatThreadsPerPage: number;
  /**
   * Method for loading more chat threads
   */
  tryLoadChatThreads(): void;
  /**
   * Method for loading more chat threads
   */
  tryLoadMoreChatThreads(): void;
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
  setSelectedChat(chatThread: GraphChatThread): void;
  /**
   * Method for setting the currently selected chat
   * @param chatId - currently selected chat id
   */
  setSelectedChatId(chatId: string): void;
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

const GraphChatThreadLastMsgPreviewCreatedComparator = (a: GraphChatThread, b: GraphChatThread): number => {
  if (a.lastMessagePreview?.createdDateTime && b.lastMessagePreview?.createdDateTime) {
    const dateA = new Date(a.lastMessagePreview.createdDateTime);
    const dateB = new Date(b.lastMessagePreview.createdDateTime);
    if (dateA === dateB) return 0;
    return dateB > dateA ? 1 : -1;
  } else if (b.lastMessagePreview?.createdDateTime) {
    return 1;
  }
  return -1;
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
  private _loadPromise: Promise<void> | undefined;
  private _loadMorePromise: Promise<void> | undefined;

  constructor() {
    this.userId = currentUserId();
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    this._cache = new LastReadCache();
    this._graph = graph('mgt-chat', GraphConfig.version);

    this._notificationClient = new GraphNotificationUserClient(this._eventEmitter, this._graph);

    this.updateUserSubscription(this.userId);
  }

  /**
   * Provides the number of chat threads to display with each load more.
   */
  public chatThreadsPerPage = 20;

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public tearDown() {
    this._notificationClient.tearDown();
  }

  private sleep(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Switches to a fatal error state and logs the error.
   */
  private raiseFatalError(e: Error) {
    error(e);
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'fatal error';
      draft.chatThreads = [];
    });
  }

  /**
   * Load more chat threads if applicable.
   */
  public async tryLoadMoreChatThreads() {
    try {
      // do not load if another activity is in progress
      if (this._loadPromise || this._loadMorePromise) {
        return;
      }

      // make sure there is something to load
      const state = this.getState();
      if (!state.moreChatThreadsToLoad) {
        return;
      }

      // set promise; load and append
      this._loadMorePromise = this.loadAndAppendChatThreads('', [], state.chatThreads.length + this.chatThreadsPerPage);
      await this._loadMorePromise;
    } finally {
      this._loadMorePromise = undefined;
    }
  }

  /**
   * Loads chat threads without regard for the existing threads.
   */
  public async tryLoadChatThreads() {
    // set the state to loading and clear all chats

    // do not load if there is another load already running
    if (this._loadPromise) {
      return;
    }

    // wait for any load more to finish
    if (this._loadMorePromise) {
      await this._loadMorePromise;
    }

    // clear and announce loading
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'loading chats';
      draft.chatThreads = [];
    });

    // try several times to load more chats
    try {
      let loaded = false;
      let count = 0;
      do {
        count++;
        try {
          this._loadPromise = this.loadAndAppendChatThreads('', [], this.chatThreadsPerPage);
          await this._loadPromise;
          loaded = true;
        } catch (e) {
          if (count > 3) {
            error('Failed to load chat threads; aborting...', e);
            throw Error('Failed to load chat threads even after 3 attempts; no more attempts will be made.');
          }
          error('Failed to load chat threads; retrying...', e);
          await this.sleep(2000);
        }
      } while (!loaded);
    } finally {
      this._loadPromise = undefined;
    }
  }

  private async handleChatThreadsResponse(
    latestChatThreads: ChatThreadCollection,
    items: GraphChatThread[],
    maxItems: number
  ) {
    const latestItems = latestChatThreads.value as GraphChatThread[];
    const checkedItems = await this.checkWhetherToMarkAsRead(latestItems);

    const idsIncheckedItems = new Set(checkedItems.map(item => item.id));
    items = items.filter(item => !idsIncheckedItems.has(item.id));
    items = items.concat(checkedItems);
    items.sort(GraphChatThreadLastMsgPreviewCreatedComparator);

    const handlerNextLink = latestChatThreads['@odata.nextLink'];

    if (items.length > maxItems) {
      // return exact page size
      this.handleChatThreads(items.slice(0, maxItems), 'more');
      return;
    }

    if (items.length < maxItems && handlerNextLink) {
      await this.loadAndAppendChatThreads(handlerNextLink, items, maxItems);
      return;
    }

    this.handleChatThreads(items, handlerNextLink);
  }

  private async loadAndAppendChatThreads(nextLink: string, items: GraphChatThread[], maxItems: number) {
    if (maxItems < 1) {
      error('maxItem is invalid: ' + maxItems);
      return;
    }

    try {
      const response = !nextLink
        ? await loadChatThreads(this._graph, maxItems > 50 ? 50 : maxItems) // max page count cannot exceed 50 per documentation
        : await loadChatThreadsByPage(this._graph, nextLink.split('?')[1]);
      await this.handleChatThreadsResponse(response, items, maxItems);
    } catch (err) {
      error(err);
    }
  }

  public clearSelectedChat = () => {
    const state = this.getState();

    if (state.internalSelectedChat) {
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'chat unselected';
        draft.internalPrevSelectedChat = state.internalSelectedChat;
        draft.internalSelectedChat = undefined;
      });
    }
  };

  public setSelectedChatId = (chatId: string) => {
    // the first time we are setting the selected chat, we may still be loading chat threads.
    // so trying the code block after this will not work as there are no chat threads yet.
    // we are setting this flag so that the `chats loaded` on ChatList can fire onselected.
    if (!this._initialSelectedChatId) {
      this._initialSelectedChatId = chatId;
      return;
    }

    const state = this.getState();
    if (!state.internalSelectedChat || state.internalSelectedChat.id !== chatId) {
      const chatThread = state.chatThreads.find(c => c.id === chatId);
      if (chatThread) {
        this.setSelectedChat(chatThread);
      }
    }
  };

  public setSelectedChat = (chatThread: GraphChatThread): void => {
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

      draft.chatThreads = state.chatThreads.map((c: GraphChatThread) => {
        if (c.id === chatThread.id && draft.internalSelectedChat) {
          return draft.internalSelectedChat;
        }
        return c;
      });

      this.updateLastReadTime(draft.internalSelectedChat.id);
    });
  };

  public markAllChatThreadsAsRead = () => {
    // mark as read after chat thread is found in current state
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'chats read';
      draft.chatThreads = this.getState().chatThreads.map((chatThread: GraphChatThread) => {
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
      const selectedChatId = state.internalSelectedChat?.id;
      this.updateLastReadTime(selectedChatId);
    }
  };

  private updateLastReadTime(selectedChatId: string | undefined) {
    if (selectedChatId) {
      log(`caching the last-read timestamp of now to chat ID '${selectedChatId}'...`);
      void this._cache.cacheLastReadTime(selectedChatId, new Date());
    }
  }

  // check whether to mark the chat as read or not
  private readonly checkWhetherToMarkAsRead = async (c: GraphChatThread[]): Promise<GraphChatThread[]> => {
    const result = await Promise.all(
      c.map(async (chatThread: GraphChatThread) => {
        // when the last message is from you, the thread is read
        if (this.userId === chatThread.lastMessagePreview?.from?.user?.id) {
          return {
            ...chatThread,
            isRead: true
          };
        }

        // when there is not a last read time, the thread is unread
        const lastReadData = await this._cache.loadLastReadTime(chatThread.id!);
        if (!lastReadData) {
          return chatThread;
        }

        // when the last message is newer than the last read time, the thread is unread
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
    fireOnSelected: false
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
        // async load the updated chat details because the notification does not include the displayName
        // of the user that was added.
        void this.loadChatDetails(chatThread.id!);
        bringToTop();
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
        } else if (this.userId === event.message.from?.user?.id) {
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
        // async load more info
        void this.loadChatDetails(event.message.chatId);
      } else {
        log(`received unrecognized event type '${event.type}' from the user subscription.`);
      }
    });
  }

  private async loadChatDetails(chatId: string) {
    try {
      const loaded = await loadChatWithPreview(this._graph, chatId);
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'chat details loaded';
        const chatThread = draft.chatThreads?.findIndex(c => c.id === chatId) ?? -1;
        if (chatThread > -1) {
          draft.chatThreads[chatThread] = Object.assign(draft.chatThreads[chatThread], loaded);
        }
      });
    } catch (e) {
      error(`Failed to load chat details for chat ID ${chatId}.`, e);
    }
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
      draft.moreChatThreadsToLoad = Boolean(nextLink);
      draft.fireOnSelected = false;
      if (this._initialSelectedChatId) {
        // again, we expect this code to only run once, during the init of ChatList component if and only if _initialSelectedChatId is set.
        const toFindId = this._initialSelectedChatId;
        const chat = chatThreads.find(c => c.id === toFindId);
        this._initialSelectedChatId = ''; // ensure we only set the selected chat once
        if (chat) {
          draft.internalSelectedChat = chat;
          draft.fireOnSelected = true;
        } else {
          log('Chat with id ' + toFindId + ' not found in loaded chat threads.');
        }
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
    this.clearCurrentUserMessages();

    // need to ensure that we close any existing connection if present
    await this._notificationClient?.closeSignalRConnection();

    // update user info
    this.userId = userId;

    // by updating the followed chat the notification client will reconnect to SignalR
    this.updateUserSubscription(userId);
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
  private updateUserSubscription(userId: string) {
    if (!userId) return;

    // reset state to initial
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'initial';
    });
    // Subscribe to notifications for messages
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'creating server connections';
    });
    try {
      this._notificationClient.subscribeToUserNotifications(userId);
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
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'server connection established';
      });
      void this.tryLoadChatThreads().catch(e => this.raiseFatalError(e as Error));
    });
  }
}

export { StatefulGraphChatListClient };
