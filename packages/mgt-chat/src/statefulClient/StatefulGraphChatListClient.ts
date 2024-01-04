import { MessageThreadProps, ErrorBarProps } from '@azure/communication-react';
import { ActiveAccountChanged, LoginChangedEvent, ProviderState, Providers } from '@microsoft/mgt-element';
import { GraphError } from '@microsoft/microsoft-graph-client';
import { produce } from 'immer';
import { currentUserId, currentUserName } from '../utils/currentUser';
import { graph } from '../utils/graph';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationUserClient } from './GraphNotificationUserClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
// defines the type of the state object returned from the StatefulGraphChatClient
export type GraphChatListClient = Pick<MessageThreadProps, 'userId'> & {
  status:
    | 'initial'
    | 'creating server connections'
    | 'subscribing to notifications'
    | 'loading messages'
    | 'no session id'
    | 'no messages'
    | 'ready'
    | 'error';
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
}

class StatefulGraphChatListClient implements StatefulClient<GraphChatListClient> {
  private readonly _notificationClient: GraphNotificationUserClient;
  private readonly _eventEmitter: ThreadEventEmitter;
  private _subscribers: ((state: GraphChatListClient) => void)[] = [];

  private _userDisplayName = '';

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this._notificationClient = new GraphNotificationUserClient(
      this._eventEmitter,
      graph('mgt-chat', GraphConfig.version)
    );
  }

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public async tearDown() {
    await this._notificationClient.tearDown();
  }

  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public onStateChange(handler: (state: GraphChatListClient) => void): void {
    if (!this._subscribers.includes(handler)) {
      this._subscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public offStateChange(handler: (state: GraphChatListClient) => void): void {
    const index = this._subscribers.indexOf(handler);
    if (index !== -1) {
      this._subscribers = this._subscribers.splice(index, 1);
    }
  }

  private readonly _initialState: GraphChatListClient = {
    status: 'initial',
    activeErrorMessages: [],
    userId: ''
  };

  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatClientList}
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
    this._subscribers.forEach(handler => handler(this._state));
  }

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  public getState(): GraphChatListClient {
    return this._state;
  }

  /**
   * Update the state of the client when the Login state changes
   *
   * @private
   * @param {LoginChangedEvent} e The event that triggered the change
   * @memberof StatefulGraphChatClient
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
    // this.clearCurrentUserMessages();
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

  private updateUserInfo() {
    this.updateCurrentUserId();
    this.updateCurrentUserName();
  }

  /**
   * Changes the userDisplayName to the current value.
   */
  private updateCurrentUserName() {
    this._userDisplayName = currentUserName();
  }

  /**
   * Changes the current user ID value to the current value.
   */
  private updateCurrentUserId() {
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

  private _sessionId: string | undefined;

  public subscribeToUser(sessionId: string) {
    this._sessionId = sessionId;
    void this.updateUserSubscription();
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatClient
   */
  private async updateUserSubscription() {
    // avoid subscribing to a resource with an empty chatId
    if (this.userId && this._sessionId) {
      // reset state to initial
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'initial';
        // draft.messages = [];
        // draft.mentions = [];
        // draft.chat = undefined;
        // draft.participants = [];
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
        tasks.push(this._notificationClient.subscribeToUserNotifications(this._userId, this._sessionId));
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
}

export { StatefulGraphChatListClient };
