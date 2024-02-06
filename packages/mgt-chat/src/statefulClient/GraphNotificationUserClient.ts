/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BetaGraph, IGraph, Providers, createFromProvider, error, log } from '@microsoft/mgt-element';
import {
  HubConnection,
  HubConnectionBuilder,
  IHttpConnectionOptions,
  LogLevel,
  RetryContext
} from '@microsoft/signalr';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import type {
  Entity,
  Subscription,
  ChatMessage,
  Chat,
  AadUserConversationMember
} from '@microsoft/microsoft-graph-types';
import { GraphConfig } from './GraphConfig';
import { SubscriptionsCache, ComponentType } from './Caching/SubscriptionCache';
import { Timer } from '../utils/Timer';

export const appSettings = {
  defaultSubscriptionLifetimeInMinutes: 10,
  renewalThreshold: 75, // The number of seconds before subscription expires it will be renewed
  renewalTimerInterval: 10, // The number of seconds between executions of the renewal timer
  useCanary: GraphConfig.useCanary
};

type ChangeTypes = 'created' | 'updated' | 'deleted';

interface Notification<T extends Entity> {
  subscriptionId: string;
  changeType: ChangeTypes;
  resource: string;
  resourceData: T & {
    id: string;
    '@odata.type': string;
    '@odata.id': string;
  };
  EncryptedContent: string;
}

type ReceivedNotification = Notification<Chat> | Notification<ChatMessage> | Notification<AadUserConversationMember>;

const isMessageNotification = (o: Notification<Entity>): o is Notification<ChatMessage> =>
  o.resource.includes('/messages(');
const isMembershipNotification = (o: Notification<Entity>): o is Notification<AadUserConversationMember> =>
  o.resource.includes('/members');

export class GraphNotificationUserClient {
  private connection?: HubConnection = undefined;
  private renewalInterval?: string;
  private renewalCount = 0;
  private isRewnewalInProgress = false;
  private userId = '';
  private currentUserId = '';
  private get sessionId() {
    return 'default';
  }
  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();
  private readonly timer = new Timer();
  private get graph() {
    return this._graph;
  }
  private get beta() {
    return BetaGraph.fromGraph(this._graph);
  }
  private get subscriptionGraph() {
    return GraphConfig.useCanary
      ? createFromProvider(Providers.globalProvider, GraphConfig.canarySubscriptionVersion, 'mgt-chat')
      : this.beta;
  }

  /**
   *
   */
  constructor(
    private readonly emitter: ThreadEventEmitter,
    private readonly _graph: IGraph
  ) {}

  /**
   * Removes any active timers that may exist to prevent memory leaks and perf issues.
   * Call this method when the component that depends an instance of this class is being removed from the DOM
   * i.e
   */
  public tearDown() {
    log('cleaning up user graph notification resources');
    if (this.renewalInterval) this.timer.clearInterval(this.renewalInterval);
    this.timer.close();
  }

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  // TODO: understand if this is needed under the native model
  private readonly onReconnect = (connectionId: string | undefined) => {
    log(`Reconnected. ConnectionId: ${connectionId || 'undefined'}`);
    const emitter: ThreadEventEmitter | undefined = this.emitter;
    emitter?.connected();
  };

  private readonly receiveNotificationMessage = (message: string) => {
    if (typeof message !== 'string') throw new Error('Expected string from receivenotificationmessageasync');

    const notification: ReceivedNotification = JSON.parse(message) as ReceivedNotification;
    log('received user notification message', notification);
    const emitter: ThreadEventEmitter | undefined = this.emitter;
    if (!notification.resourceData) throw new Error('Message did not contain resourceData');
    if (isMessageNotification(notification)) {
      this.processMessageNotification(notification, emitter);
    } else if (isMembershipNotification(notification)) {
      this.processMembershipNotification(notification, emitter);
    } else {
      this.processChatPropertiesNotification(notification, emitter);
    }
    // Need to return a status code string of 200 so that graph knows the message was received and doesn't re-send the notification
    const ackMessage: unknown = { StatusCode: '200' };
    return GraphConfig.ackAsString ? JSON.stringify(ackMessage) : ackMessage;
  };

  private processMessageNotification(notification: Notification<ChatMessage>, emitter: ThreadEventEmitter | undefined) {
    const message = notification.resourceData;

    switch (notification.changeType) {
      case 'created':
        emitter?.chatMessageReceived(message);
        return;
      case 'updated':
        emitter?.chatMessageEdited(message);
        return;
      case 'deleted':
        emitter?.chatMessageDeleted(message);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private processMembershipNotification(
    notification: Notification<AadUserConversationMember>,
    emitter: ThreadEventEmitter | undefined
  ) {
    const member = notification.resourceData;
    switch (notification.changeType) {
      case 'created':
        emitter?.participantAdded(member);
        return;
      case 'deleted':
        emitter?.participantRemoved(member);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private processChatPropertiesNotification(notification: Notification<Chat>, emitter: ThreadEventEmitter | undefined) {
    const chat = notification.resourceData;
    switch (notification.changeType) {
      case 'updated':
        emitter?.chatThreadPropertiesUpdated(chat);
        return;
      case 'deleted':
        emitter?.chatThreadDeleted(chat);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private readonly cacheSubscription = async (userId: string, subscriptionRecord: Subscription): Promise<void> => {
    log(subscriptionRecord);
    await this.subscriptionCache.cacheSubscription(userId, ComponentType.User, this.sessionId, subscriptionRecord);
  };

  private async subscribeToResource(userId: string, resourcePath: string, changeTypes: ChangeTypes[]) {
    // build subscription request
    const expirationDateTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    ).toISOString();
    const subscriptionDefinition: Subscription = {
      changeType: changeTypes.join(','),
      notificationUrl: `${GraphConfig.webSocketsPrefix}?groupId=${userId}&sessionId=${this.sessionId}`,
      resource: resourcePath,
      expirationDateTime,
      includeResourceData: true,
      clientState: 'wsssecret'
    };

    log('subscribing to changes for ' + resourcePath);
    const subscriptionEndpoint = GraphConfig.subscriptionEndpoint;
    // send subscription POST to Graph
    const subscription: Subscription = (await this.subscriptionGraph
      .api(subscriptionEndpoint)
      .post(subscriptionDefinition)) as Subscription;
    if (!subscription?.notificationUrl) throw new Error('Subscription not created');
    log(subscription);

    await this.cacheSubscription(userId, subscription);

    // create a connection to the web socket if one does not exist
    if (!this.connection) await this.createSignalRConnection(subscription.notificationUrl);

    log('Invoked CreateSubscription');
  }

  private readonly renewalSync = () => {
    void this.renewal();
  };

  private readonly renewal = async () => {
    if (this.isRewnewalInProgress) {
      log('Renewal already in progress');
      return;
    }

    this.isRewnewalInProgress = true;

    this.currentUserId = this.userId;

    try {
      const subscriptions =
        (await this.subscriptionCache.loadSubscriptions(this.currentUserId, this.sessionId))?.subscriptions || [];
      if (subscriptions.length === 0) {
        log('No subscriptions found in session state. Creating a new subscription.');

        await this.subscribeToResource(this.currentUserId, `/users/${this.currentUserId}/chats/getAllmessages`, [
          'created',
          'updated',
          'deleted'
        ]);
      } else {
        for (const subscription of subscriptions) {
          if (!subscription.expirationDateTime || !subscription.id || !subscription.notificationUrl) continue;

          const expirationTime = new Date(subscription.expirationDateTime);
          const now = new Date();
          const diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

          if (diff <= appSettings.renewalThreshold) {
            this.renewalCount++;
            log(`Renewing Graph subscription for ChatList. RenewalCount: ${this.renewalCount}.`);

            const newExpirationTime = new Date(
              new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
            );

            try {
              await this.renewSubscription(this.currentUserId, subscription.id, newExpirationTime.toISOString());
            } catch (e) {
              // this error indicates we are not able to successfully renew the subscription, so we should create a new one.
              if ((e as { statusCode?: number }).statusCode === 404) {
                log('Removing subscription from cache', subscription.id);
                await this.subscriptionCache.deleteCachedSubscriptions(this.currentUserId, this.sessionId);
                await this.subscribeToUserNotifications(this.currentUserId);

                const emitter: ThreadEventEmitter | undefined = this.emitter;
                emitter?.reconnected();
              }
            }
          } else {
            // create a connection to the web socket if one does not exist
            if (!this.connection) await this.createSignalRConnection(subscription.notificationUrl);
          }

          // Expecting only one subscription per user
          break;
        }
      }
    } catch (e) {
      error(e);
    }

    this.isRewnewalInProgress = false;
    this.renewalInterval = this.timer.setTimeout(this.renewalSync, appSettings.renewalTimerInterval * 1000);
  };

  private readonly renewSubscription = async (
    userId: string,
    subscriptionId: string,
    expirationDateTime: string
  ): Promise<void> => {
    // PATCH /subscriptions/{id}
    const renewedSubscription = (await this.graph.api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`).patch({
      expirationDateTime
    })) as Subscription;
    return this.cacheSubscription(userId, renewedSubscription);
  };

  private async createSignalRConnection(notificationUrl: string) {
    log('Creating SignalR connection');

    const connectionOptions: IHttpConnectionOptions = {
      accessTokenFactory: this.getToken,
      withCredentials: false
    };

    // retry with the following intervals, the last interval will take precedence if there are more retries than intervals
    const retryTimes = [0, 2000, 10000, 30000, 45000, 60000, 90000, 120000, 180000, 240000];
    const retryPolicy = {
      nextRetryDelayInMilliseconds: (context: RetryContext) => {
        const index =
          context.previousRetryCount < retryTimes.length ? context.previousRetryCount : retryTimes.length - 1;
        return retryTimes[index];
      }
    };

    const connection = new HubConnectionBuilder()
      .withUrl(GraphConfig.adjustNotificationUrl(notificationUrl), connectionOptions)
      .withAutomaticReconnect(retryPolicy)
      .configureLogging(LogLevel.Information)
      .build();

    const emitter: ThreadEventEmitter | undefined = this.emitter;
    connection.onclose((err?: Error) => {
      if (err) {
        log('Connection closed with error', err);
      }

      emitter?.disconnected();
    });

    connection.onreconnected(this.onReconnect);

    connection.onreconnecting(() => {
      emitter?.disconnected();
    });

    connection.on('receivenotificationmessageasync', this.receiveNotificationMessage);

    connection.on('EchoMessage', log);

    this.connection = connection;
    try {
      await connection.start();
      log(connection);
      emitter?.connected();
    } catch (e) {
      error('An error occurred connecting to the notification web socket', e);
    }
  }

  private async deleteSubscription(id: string) {
    try {
      await this.graph.api(`${GraphConfig.subscriptionEndpoint}/${id}`).delete();
      log(`Deleted subscription with id: ${id}`);
    } catch (e) {
      error(e);
    }
  }

  private async removeSubscriptions(subscriptions: Subscription[]): Promise<unknown[]> {
    const tasks: Promise<unknown>[] = [];
    for (const s of subscriptions) {
      // if there is no id or the subscription is expired, skip
      if (!s.id || (s.expirationDateTime && new Date(s.expirationDateTime) <= new Date())) continue;
      tasks.push(this.deleteSubscription(s.id));
    }
    return Promise.all(tasks);
  }

  public async closeSignalRConnection() {
    // stop the connection and set it to undefined so it will reconnect when next subscription is created.
    await this.connection?.stop();
    this.connection = undefined;
  }

  public async unsubscribeFromUserNotifications(userId: string) {
    if (this.renewalInterval) {
      this.timer.clearTimeout(this.renewalInterval);
    }

    await this.closeSignalRConnection();
    const cacheData = await this.subscriptionCache.loadSubscriptions(userId, this.sessionId);
    if (cacheData) {
      await Promise.all([
        this.removeSubscriptions(cacheData.subscriptions),
        this.subscriptionCache.deleteCachedSubscriptions(userId, this.sessionId)
      ]);
    }
  }

  public async subscribeToUserNotifications(userId: string) {
    log(`User subscription with id: ${userId}`);
    this.userId = userId;
    await this.renewal();
  }
}
