import { ErrorBarProps } from '@azure/communication-react';
import { graph } from '../utils/graph';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationClient } from './GraphNotificationClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
// defines the type of the state object returned from the StatefulGraphChatClient
export type GraphChatListClient = Pick<ErrorBarProps, 'activeErrorMessages'>;

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
  private readonly _notificationClient: GraphNotificationClient;
  private readonly _eventEmitter: ThreadEventEmitter;
  private _subscribers: ((state: GraphChatListClient) => void)[] = [];

  constructor() {
    this._eventEmitter = new ThreadEventEmitter();
    this._notificationClient = new GraphNotificationClient(this._eventEmitter, graph('mgt-chat', GraphConfig.version));
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
    activeErrorMessages: []
  };

  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatClientList}
   * @memberof StatefulGraphChatListClient
   */
  private readonly _state: GraphChatListClient = { ...this._initialState };

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  public getState(): GraphChatListClient {
    return this._state;
  }
}

export { StatefulGraphChatListClient };
