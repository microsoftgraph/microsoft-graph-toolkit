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
  HubConnectionState,
  IHttpConnectionOptions,
  LogLevel
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
  renewalTimerInterval: 3, // The number of seconds between executions of the renewal timer
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
  private wasConnected?: boolean | undefined;
  private userId = '';
  private currentUserId = '';
  private lastNotificationUrl = '';

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
    void this.unsubscribeFromUserNotifications(this.userId);
  }

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
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
    await this.subscriptionCache.cacheSubscription(userId, ComponentType.User, subscriptionRecord);
  };

  private async createSubscription(userId: string): Promise<Subscription> {
    log('Creating a new subscription.');
    const resourcePath = `/users/${userId}/chats/getAllmessages`;
    const changeTypes: ChangeTypes[] = ['created', 'updated', 'deleted'];

    // build subscription request
    const expirationDateTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    ).toISOString();
    const subscriptionDefinition: Subscription = {
      changeType: changeTypes.join(','),
      notificationUrl: `${GraphConfig.webSocketsPrefix}?groupId=${userId}&sessionId=default`,
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

    log('Subscription created.');

    return subscription;
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

    let isRenewalInError = false;

    try {
      let subscription = await this.getSubscription(this.currentUserId);

      if (subscription) {
        if (!subscription.expirationDateTime || !subscription.id) {
          // this should never happen.
          throw new Error('Subscription is invalid.');
        }

        try {
          const expirationTime = new Date(subscription.expirationDateTime);
          const now = new Date();
          const diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

          if (diff <= appSettings.renewalThreshold) {
            await this.renewSubscription(this.currentUserId, subscription);
          }
        } catch (renewalEx) {
          isRenewalInError = true;
          // this error indicates we are not able to successfully renew the subscription, so we should create a new one.
          if ((renewalEx as { statusCode?: number }).statusCode === 404) {
            log('Removing subscription from cache');
            await this.subscriptionCache.deleteCachedSubscriptions(this.currentUserId);
            subscription = undefined;
          } else {
            // log and continue (we will try again later)
            error(renewalEx);
          }
        }
      }

      if (!subscription) {
        subscription = await this.createSubscription(this.currentUserId);
      }

      // notificationUrl comes in the form of websockets:https://graph.microsoft.com/beta/subscriptions/notificationChannel/websockets/<Id>?groupid=<UserId>&sessionid=default
      // if <Id> changes, we need to create a new connection
      if (
        !this.connection ||
        this.connection.state !== HubConnectionState.Connected ||
        this.lastNotificationUrl !== subscription.notificationUrl
      ) {
        if (subscription.notificationUrl) {
          this.lastNotificationUrl = subscription.notificationUrl;
          await this.createSignalRConnection(subscription.notificationUrl);
        } else {
          // this should never happen.
          throw new Error('Subscription notificationUrl is invalid.');
        }
      }
    } catch (e) {
      isRenewalInError = true;
      // log and continue (we will try again later)
      error(e);
    }

    const isConnected = !isRenewalInError && this.connection?.state === HubConnectionState.Connected;
    if (this.wasConnected !== isConnected) {
      this.wasConnected = isConnected;
      const emitter: ThreadEventEmitter | undefined = this.emitter;

      try {
        if (isConnected) {
          emitter?.connected();
        } else {
          emitter?.disconnected();
        }
      } catch (e) {
        // log emitter thrown exception and continue
        error(e);
      }
    }

    this.isRewnewalInProgress = false;
    this.renewalInterval = this.timer.setTimeout(this.renewalSync, appSettings.renewalTimerInterval * 1000);
  };

  private async getSubscription(userId: string): Promise<Subscription | undefined> {
    const subscriptions = (await this.subscriptionCache.loadSubscriptions(userId))?.subscriptions || [];
    return subscriptions.length > 0 ? subscriptions[0] : undefined;
  }

  private readonly renewSubscription = async (userId: string, subscription: Subscription): Promise<void> => {
    this.renewalCount++;
    log(`Renewing Graph subscription for ChatList. RenewalCount: ${this.renewalCount}.`);

    const newExpirationTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );
    // PATCH /subscriptions/{id}
    const subscriptionId = subscription.id;
    const expirationDateTime = newExpirationTime.toISOString();
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

    const connection = new HubConnectionBuilder()
      .withUrl(GraphConfig.adjustNotificationUrl(notificationUrl), connectionOptions)
      .configureLogging(LogLevel.Information)
      .build();

    connection.on('receivenotificationmessageasync', this.receiveNotificationMessage);
    connection.on('EchoMessage', log);

    this.connection = connection;
    try {
      await connection.start();
      log(connection);
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
    const cacheData = await this.subscriptionCache.loadSubscriptions(userId);
    if (cacheData) {
      await Promise.all([
        this.removeSubscriptions(cacheData.subscriptions),
        this.subscriptionCache.deleteCachedSubscriptions(userId)
      ]);
    }
  }

  public async subscribeToUserNotifications(userId: string) {
    log(`User subscription with id: ${userId}`);
    this.wasConnected = undefined;
    this.userId = userId;
    await this.renewal();
  }
}
