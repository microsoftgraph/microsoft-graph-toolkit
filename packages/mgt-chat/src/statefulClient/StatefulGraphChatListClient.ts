import { MessageThreadProps, ErrorBarProps, Message } from '@azure/communication-react';
import { ActiveAccountChanged, IGraph, LoginChangedEvent, ProviderState, Providers, log } from '@microsoft/mgt-element';
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
// TODO: MessageCache is added here for the purpose of following the convention of StatefulGraphChatClient. However, StatefulGraphChatListClient
//       is also leveraging the same cache and performing the same actions which would have resulted in race conditions against the same messages
//       in the cache. To avoid this, I have commented out the code. We should revisit this and determine if we need to use the cache.
// import { MessageCache } from './Caching/MessageCache';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationUserClient } from './GraphNotificationUserClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { ChatThreadCollection, loadChatThreads, loadChatThreadsByPage } from './graph.chat';
import { ChatMessageInfo, Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { error } from '@microsoft/mgt-element';
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

// defines the type of the state object returned from the StatefulGraphChatListClient
export type GraphChatListClient = Pick<MessageThreadProps, 'userId'> & {
  status:
    | 'initial'
    | 'creating server connections'
    | 'subscribing to notifications'
    | 'loading messages'
    | 'no session id'
    | 'no messages'
    | 'chat threads loaded'
    | 'ready'
    | 'error';
  chatThreads: GraphChat[];
  moreChatThreadsToLoad: boolean | undefined;
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
   * Register a callback to receive ChatList events
   *
   * @param handler Callback to receive ChatList events
   */
  onChatListEvent(handler: (event: ChatListEvent) => void): void;
  /**
   * Remove a callback to receive ChatList events
   *
   * @param handler Callback to be unregistered
   */
  offChatListEvent(handler: (event: ChatListEvent) => void): void;

  chatThreadsPerPage: number;

  loadMoreChatThreads(): void;
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
  // private readonly _cache: MessageCache;
  private _stateSubscribers: ((state: GraphChatListClient) => void)[] = [];
  private _chatListEventSubscribers: ((state: ChatListEvent) => void)[] = [];
  private readonly _graph: IGraph;
  constructor(chatThreadsPerPage: number) {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    // this._cache = new MessageCache();
    this._graph = graph('mgt-chat', GraphConfig.version);
    this.chatThreadsPerPage = chatThreadsPerPage;
    this._notificationClient = new GraphNotificationUserClient(this._eventEmitter, this._graph);
  }

  /**
   * Provides the number of chat threads to display with each load more.
   */
  public chatThreadsPerPage: number;

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public async tearDown() {
    await this._notificationClient.tearDown();
  }

  /**
   * Load more chat threads if applicable.
   */
  public loadMoreChatThreads(): void {
    const state = this.getState();
    const items: GraphChat[] = [];
    this.loadAndAppendChatThreads('', items, state.chatThreads.length + this.chatThreadsPerPage);
  }

  private loadAndAppendChatThreads(nextLink: string, items: GraphChat[], maxItems: number): void {
    if (maxItems < 1) {
      error('maxItem is invalid: ' + maxItems);
      return;
    }

    const handler = (latestChatThreads: ChatThreadCollection) => {
      items = items.concat(latestChatThreads.value);

      const handlerNextLink = latestChatThreads['@odata.nextLink'];
      if (items.length >= maxItems) {
        if (items.length > maxItems) {
          // return exact page size
          this.handleChatThreads(items.slice(0, maxItems), 'more');
          return;
        }

        this.handleChatThreads(items, handlerNextLink);
        return;
      }

      if (handlerNextLink && handlerNextLink !== '') {
        this.loadAndAppendChatThreads(handlerNextLink, items, maxItems);
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
      loadChatThreadsByPage(this._graph, filter).then(handler, err => error(err));
    }
  }

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

  /**
   * Register a callback to receive ChatList events
   *
   * @param {(event: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public onChatListEvent(handler: (event: ChatListEvent) => void): void {
    if (!this._chatListEventSubscribers.includes(handler)) {
      this._chatListEventSubscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving ChatList events
   *
   * @param {(event: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public offChatListEvent(handler: (event: ChatListEvent) => void): void {
    const index = this._chatListEventSubscribers.indexOf(handler);
    if (index !== -1) {
      this._chatListEventSubscribers = this._chatListEventSubscribers.splice(index, 1);
    }
  }

  private readonly _initialState: GraphChatListClient = {
    status: 'initial',
    activeErrorMessages: [],
    userId: '',
    chatThreads: [],
    moreChatThreadsToLoad: undefined
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
      const chatThread = draft.chatThreads[chatThreadIndex];

      // func to bring the chat thread to the top of the list
      const bringToTop = (newThread?: GraphChat) => {
        draft.chatThreads.splice(chatThreadIndex, 1);
        draft.chatThreads.unshift(newThread ?? chatThread);
      };

      // handle the events
      if (event.type === 'chatRenamed' && event.message?.eventDetail && chatThread) {
        // rename the chat thread in-place
        chatThread.topic =
          (event.message.eventDetail as ChatRenamedEventMessageDetail)?.chatDisplayName ?? chatThread.topic;
      } else if (event.type === 'memberAdded' && chatThread) {
        // inject a chat thread with only the id to force a reload; this is necessary because the
        // notification does not include the displayName of the user who was added
        const newThread = { id: chatThread.id } as GraphChat;
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
        // update the last message preview and bring to the top
        chatThread.lastMessagePreview = event.message as ChatMessageInfo;
        chatThread.lastUpdatedDateTime = event.message.lastModifiedDateTime;
        bringToTop();
      } else if (event.type === 'chatMessageReceived' && event.message?.chatId) {
        // create a new chat thread at the top
        const newChatThread: GraphChat = {
          id: event.message.chatId,
          lastMessagePreview: event.message as ChatMessageInfo,
          lastUpdatedDateTime: event.message.lastModifiedDateTime
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
  private readonly handleChatThreads = (chatThreads: GraphChat[], nextLink: string | undefined) => {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'chat threads loaded';
      draft.chatThreads = chatThreads;
      draft.moreChatThreadsToLoad = nextLink !== undefined && nextLink !== '';
    });
  };

  /**
   * Update the state of the client when the Login state changes
   *
   * @private
   * @param {LoginChangedEvent} e The event that triggered the change
   * @memberof StatefulGraphChatListClient
   */
  private readonly onLoginStateChanged = (e: LoginChangedEvent) => {
    switch (e.detail) {
      case ProviderState.SignedIn:
        // update userId and displayName
        this.updateUserInfo();
        // load messages?
        // configure subscriptions
        // emit new state;
        if (this.userId) {
          void this.updateUserSubscription();
          this.loadAndAppendChatThreads('', [], this.chatThreadsPerPage);
        }
        return;
      case ProviderState.SignedOut:
        // clear userId
        // clear subscriptions
        // clear messages
        // emit new state
        return;
      case ProviderState.Loading:
      default:
        // do nothing for now
        return;
    }
  };

  private readonly onActiveAccountChanged = (e: ActiveAccountChanged) => {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    if (e.detail && this.userId !== e.detail?.id) {
      void this.handleAccountChange();
    }
  };

  private readonly handleAccountChange = async () => {
    this.clearCurrentUserMessages();
    // need to ensure that we close any existing connection if present
    await this._notificationClient?.closeSignalRConnection();

    this.updateUserInfo();
    // by updating the followed chat the notification client will reconnect to SignalR
    await this.updateUserSubscription();
  };

  private clearCurrentUserMessages() {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'initial'; // no message?
    });
  }

  /**
   * Changes the current user ID value to the current value.
   */
  private updateUserInfo() {
    this.userId = currentUserId();
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

  public get sessionId(): string {
    return 'default';
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatListClient
   */
  private async updateUserSubscription() {
    // avoid subscribing to a resource with an empty userId
    if (this.userId) {
      // reset state to initial
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'initial';
      });
      // Subscribe to notifications for messages
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'creating server connections';
      });
      try {
        // Prefer sequential promise resolving to catch loading message errors
        // TODO: in parallel promise resolving, find out how to trigger different
        // TODO: state for failed subscriptions in GraphChatClient.onSubscribeFailed
        const tasks: Promise<unknown>[] = [];
        // subscribing to notifications will trigger the chatMessageNotificationsSubscribed event
        // this client will then load the chat and messages when that event listener is called
        tasks.push(this._notificationClient.subscribeToUserNotifications(this._userId, this.sessionId));
        await Promise.all(tasks);
      } catch (e) {
        console.error('Failed to load chat data or subscribe to notications: ', e);
        if (e instanceof GraphError) {
          this.notifyStateChange((draft: GraphChatListClient) => {
            draft.status = 'no messages';
          });
        }
      }
    } else {
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'no session id';
      });
    }
  }

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', (message: ChatMessage) => void this.onMessageReceived(message));
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', (message: ChatMessage) => void this.onMessageEdited(message));
  }
}

export { StatefulGraphChatListClient };
