/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BetaGraph, IGraph, Providers, error, log } from '@microsoft/mgt-element';
import { HubConnection, HubConnectionBuilder, IHttpConnectionOptions, LogLevel } from '@microsoft/signalr';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import type {
  Entity,
  Subscription,
  ChatMessage,
  Chat,
  AadUserConversationMember
} from '@microsoft/microsoft-graph-types';
import { GraphConfig } from './GraphConfig';
import { SubscriptionsCache } from './Caching/SubscriptionCache';

export const appSettings = {
  defaultSubscriptionLifetimeInMinutes: 10,
  renewalThreshold: 75, // The number of seconds before subscription expires it will be renewed
  timerInterval: 20 // The number of seconds the timer will check on expiration
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

const stripWssScheme = (notificationUrl: string): string => notificationUrl.replace('websockets:', '');

export class GraphNotificationClient {
  private connection?: HubConnection = undefined;
  private renewalInterval = -1;
  private renewalCount = 0;
  private currentUserId = '';
  private chatId = '';
  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();
  private get graph() {
    return this._graph;
  }
  private get beta() {
    return BetaGraph.fromGraph(this._graph);
  }

  /**
   *
   */
  constructor(
    private readonly emitter: ThreadEventEmitter,
    private readonly _graph: IGraph
  ) {}

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  // TODO: understand if this is needed under the native model
  private readonly onReconnect = (connectionId: string | undefined) => {
    log(`Reconnected. ConnectionId: ${connectionId || 'undefined'}`);
    // void this.renewChatSubscriptions();
  };

  private async decryptMessage<T>(encryptedData: string): Promise<T> {
    return (await Promise.resolve(encryptedData)) as T;
  }

  private readonly receiveNotificationMessage = (message: string) => {
    if (typeof message !== 'string') throw new Error('Expected string from receivenotificationmessageasync');

    const notification: ReceivedNotification = JSON.parse(message) as ReceivedNotification;
    log('received notification message', notification);
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
    return { StatusCode: '200' };
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

  private readonly cacheSubscription = async (subscriptionRecord: Subscription): Promise<void> => {
    log(subscriptionRecord);

    await this.subscriptionCache.cacheSubscription(this.chatId, subscriptionRecord);

    // only start timer once. -1 for renewalInterval is semaphore it has stopped.
    if (this.renewalInterval === -1) this.startRenewalTimer();
  };

  private async subscribeToResource(resourcePath: string, changeTypes: ChangeTypes[]) {
    // build subscription request
    const expirationDateTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    ).toISOString();
    const subscriptionDefinition: Subscription = {
      changeType: changeTypes.join(','),
      notificationUrl: 'websockets:',
      resource: resourcePath,
      expirationDateTime,
      includeResourceData: true,
      clientState: 'wsssecret'
    };

    log('subscribing to changes for ' + resourcePath);
    // send subscription POST to Graph
    const subscription: Subscription = (await this.beta
      .api(GraphConfig.subscriptionEndpoint)
      .post(subscriptionDefinition)) as Subscription;
    if (!subscription?.notificationUrl) throw new Error('Subscription not created');
    log(subscription);
    // subscription.notificationUrl = GraphConfig.adjustNotificationUrl(subscription.notificationUrl);

    const awaits: Promise<void>[] = [];
    // Cache the subscription in storage for re-hydration on page refreshes
    awaits.push(this.cacheSubscription(subscription));

    // create a connection to the web socket if one does not exist
    if (!this.connection) awaits.push(this.createSignalRConnection(subscription.notificationUrl));

    log('Invoked CreateSubscription');
    return Promise.all(awaits);
  }

  private readonly startRenewalTimer = () => {
    if (this.renewalInterval !== -1) clearInterval(this.renewalInterval);
    this.renewalInterval = window.setInterval(this.syncTimerWrapper, appSettings.timerInterval * 1000);
    log(`Start renewal timer . Id: ${this.renewalInterval}`);
  };

  private readonly syncTimerWrapper = () => void this.renewalTimer();

  private readonly renewalTimer = async () => {
    const subscriptions = (await this.subscriptionCache.loadSubscriptions())?.subscriptions || [];
    if (subscriptions.length === 0) {
      log(`No subscriptions found in session state. Stop renewal timer ${this.renewalInterval}.`);
      clearInterval(this.renewalInterval);
      return;
    }

    for (const subscription of subscriptions) {
      if (!subscription.expirationDateTime) continue;
      const expirationTime = new Date(subscription.expirationDateTime);
      const now = new Date();
      const diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

      if (diff <= appSettings.renewalThreshold) {
        this.renewalCount++;
        log(`Renewing Graph subscription. RenewalCount: ${this.renewalCount}`);
        // stop interval to prevent new invokes until refresh is ready.
        clearInterval(this.renewalInterval);
        this.renewalInterval = -1;
        void this.renewChatSubscriptions();
        // There is one subscription that need expiration, all subscriptions will be renewed
        break;
      }
    }
  };

  public renewChatSubscriptions = async () => {
    clearInterval(this.renewalInterval);

    const expirationTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );

    const subscriptionCache = await this.subscriptionCache.loadSubscriptions();
    const awaits: Promise<unknown>[] = [];
    for (const subscription of subscriptionCache?.subscriptions || []) {
      if (!subscription.id) continue;
      // the renewSubscription method caches the updated subscription to track the new expiration time
      awaits.push(this.renewSubscription(subscription.id, expirationTime.toISOString()));
      log(`Invoked RenewSubscription ${subscription.id}`);
    }
    await Promise.all(awaits);
  };

  public renewSubscription = async (subscriptionId: string, expirationDateTime: string): Promise<void> => {
    // PATCH /subscriptions/{id}
    const renewedSubscription = (await this.graph.api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`).patch({
      expirationDateTime
    })) as Subscription;
    return this.cacheSubscription(renewedSubscription);
  };

  public async createSignalRConnection(notificationUrl: string) {
    const connectionOptions: IHttpConnectionOptions = {
      accessTokenFactory: this.getToken,
      withCredentials: false
    };
    const connection = new HubConnectionBuilder()
      .withUrl(stripWssScheme(notificationUrl), connectionOptions)
      .withAutomaticReconnect()
      .configureLogging(LogLevel.Information)
      .build();

    connection.onreconnected(this.onReconnect);

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

  private async removeSubscriptions(subscriptions: Subscription[]): Promise<unknown[]> {
    const tasks: Promise<unknown>[] = [];
    for (const s of subscriptions) {
      // if there is no id or the subscription is expired, skip
      if (!s.id || (s.expirationDateTime && new Date(s.expirationDateTime) <= new Date())) continue;
      tasks.push(this.graph.api(`${GraphConfig.subscriptionEndpoint}/${s.id}`).delete());
    }
    return Promise.all(tasks);
  }

  public async closeSignalRConnection() {
    // stop the connection and set it to undefined so it will reconnect when next subscription is created.
    await this.connection?.stop();
    this.connection = undefined;
  }

  public async subscribeToChatNotifications(userId: string, threadId: string, afterConnection: () => void) {
    afterConnection();

    this.currentUserId = userId;
    this.chatId = threadId;
    // MGT uses a per-user cache, so no concerns of loading the cachced data for another user.
    const cacheData = await this.subscriptionCache.loadSubscriptions();
    if (cacheData && cacheData.chatId !== threadId) {
      // check validity of subscription and delete if still valid;
      await this.removeSubscriptions(cacheData.subscriptions);
    }
    if (cacheData && cacheData.chatId === threadId) {
      // check subscription validity & renew if all still valid otherwise recreate
      const someExpired = cacheData.subscriptions.some(
        s => s.expirationDateTime && new Date(s.expirationDateTime) <= new Date()
      );
      // for a given user they only get one websocket and receive all notifications via that websocket.
      const webSocketUrl = cacheData.subscriptions.find(s => s.notificationUrl)?.notificationUrl;
      if (someExpired) {
        await this.removeSubscriptions(cacheData.subscriptions);
      } else if (webSocketUrl) {
        await this.createSignalRConnection(webSocketUrl);
        await this.renewChatSubscriptions();
        return;
      }
      await this.subscriptionCache.clearCachedSubscriptions();
    }
    const promises: Promise<unknown>[] = [];
    promises.push(this.subscribeToResource(`/chats/${threadId}/messages`, ['created', 'updated', 'deleted']));
    promises.push(this.subscribeToResource(`/chats/${threadId}/members`, ['created', 'deleted']));
    promises.push(this.subscribeToResource(`/chats/${threadId}`, ['updated', 'deleted']));
    await Promise.all(promises);
  }
}
