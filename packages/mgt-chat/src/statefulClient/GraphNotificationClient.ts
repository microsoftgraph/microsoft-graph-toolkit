/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Providers, error, log } from '@microsoft/mgt-element';
import * as signalR from '@microsoft/signalr';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import type { Subscription, ChatMessage, Chat, AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import {} from '@microsoft/microsoft-graph-types-beta';
import { GraphConfig } from './GraphConfig';
import { SubscriptionsCache } from './Caching/SubscriptionCache';

export const appSettings = {
  defaultSubscriptionLifetimeInMinutes: 10,
  renewalThreshold: 75, // The number of seconds before subscription expires it will be renewed
  timerInterval: 20 // The number of seconds the timer will check on expiration
};

type ChangeTypes = 'created' | 'updated' | 'deleted';

type Notification = {
  subscriptionId: string;
  changeType: ChangeTypes;
  resource: string;
  resourceData: {
    id: string;
    '@odata.type': string;
    '@odata.id': string;
  };
  encryptedContent: string;
};

export class GraphNotificationClient {
  private connection?: signalR.HubConnection = undefined;
  private renewalInterval = -1;
  private renewalCount = 0;
  private currentUserId = '';
  private chatId = '';
  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();

  /**
   *
   */
  constructor(private readonly emitter: ThreadEventEmitter) {}

  private get _publicKey() {
    return 'LS0tLS1CRUdJTiBDRVJUSUZJQ0FURS0tLS0tDQpNSUlGblRDQ0E0V2dBd0lCQWdJVUdJL1N3SDhjMDZ0NzB3VWtrTUxnN2UrN2hha3dEUVlKS29aSWh2Y05BUUVMDQpCUUF3WGpFTE1Ba0dBMVVFQmhNQ2RYTXhDekFKQmdOVkJBZ01BbmRoTVJBd0RnWURWUVFIREFkeVpXUnRiMjVrDQpNUkF3RGdZRFZRUUtEQWRqYjI1MGIzTnZNUXd3Q2dZRFZRUUxEQU5rWlhZeEVEQU9CZ05WQkFNTUIyMW5kQzVrDQpaWFl3SGhjTk1qTXdPREUzTVRnMU5UVTRXaGNOTWpRd09ERTJNVGcxTlRVNFdqQmVNUXN3Q1FZRFZRUUdFd0oxDQpjekVMTUFrR0ExVUVDQXdDZDJFeEVEQU9CZ05WQkFjTUIzSmxaRzF2Ym1ReEVEQU9CZ05WQkFvTUIyTnZiblJ2DQpjMjh4RERBS0JnTlZCQXNNQTJSbGRqRVFNQTRHQTFVRUF3d0hiV2QwTG1SbGRqQ0NBaUl3RFFZSktvWklodmNODQpBUUVCQlFBRGdnSVBBRENDQWdvQ2dnSUJBTFEyQnVKNWdwZ3RjMTYwM0hVMlMrUjBJaFlqOHRNd29UQ1FCc2pVDQpxVW44ZS9GRDduaUQ2ZGRpNWVvRzdXZkdHd2MrUnIzS0tYV3VDemJRQlJnb0xLZk8wbUdtVWFuaEt1a3JKYXBqDQpxYmoycDZReERYMTJCQjlORHVrQ1NEZy9ZdmdjeThTRHBEYTBSL3pyTE9UU2VTb0J1MzhzbGNEbmxBcDVMbTc3DQorTVVqNFV2cG5lWFIrOXRFUjFBQWQySVpZT1RTTFM2bllId2plUndtQ2FhV1VITHYrdnR2emQ3MUdiWmlUenBRDQpDNXNtS1dJUmVzM3VGOGoyb3hXcndJY042WVZuV3lQb3RXT2NNZEhSdmdyVDZSSjdQSCtkYmlMUlJ5OW1ORk94DQptbUJJY2ZxaWRSVHREZlFEYndNeExhNXArRlNvRW92QWhNUi9rUFRMaDkzSkkyZDJxVStEQnJnQkhYZ0dRSHZDDQo4TUZsaVBpVGJwZWJ3R09sbkNHQnZpdWp2cEJ2TUJscUhJU3RlTE0zOGwvUFRna1VsVmtrMURrdjgvRko5TlZHDQpZN2piUmV5ck4wbmlpcXRadHNUOVpNb1FQNnErdWpON0M1clA5NzdVQzlySnZ6TG9TcVhIVVh3SWhRVXUrd3dPDQpnYlBmTVpyZG1aT1NGRzMxSjQvT0tDRW0zMGNtZFR3VmdJeUJOUjNBTlJXUDdtVzdBakFsSmY5ckszVkZMK3FDDQpUM244RXlJcXVQRnVDd05uWGpWbXpQNzJZRU5MWEE4Y0lYczNFVGNDeE1VQjZ1RTlJY2hOK1ROemdpeklsNTVSDQpKa0RETFdoVmFlcHUzOFRpcldWTHI5TG8wZWRwemRMcmZoOHNXU3JJbjRYaVRTWE94d2VEMUNldW42MkRCSFl0DQpGSnMzQWdNQkFBR2pVekJSTUIwR0ExVWREZ1FXQkJScmhPQ0pZNUFUeWNTNkdhT3BwS3NscnlUWW5qQWZCZ05WDQpIU01FR0RBV2dCUnJoT0NKWTVBVHljUzZHYU9wcEtzbHJ5VFluakFQQmdOVkhSTUJBZjhFQlRBREFRSC9NQTBHDQpDU3FHU0liM0RRRUJDd1VBQTRJQ0FRQi8wRTZnd0lpMU1ESHkxc1hISFJQb2tYbVdkeFZ0YmY3YUw3elI3YllkDQpKaUVXRzEzc1J6SElFa3pETTMwUkRJT2xvZnRrbUM1bUM2R09XMmh2WTk4alMzQnM4Z0xIVktwNTU2NVpoT01mDQpheWVId0ZmK1NON1VMRDhMYVNsSTNqVkIydzlkZFFObjhHbm1oejBYckJqektOcEl0NzVyYXFaSGpXRGVwR3lxDQpkS1dsSXJ6RktSOVF6dVV4MzhTeDNxSzlaMWhRdHBQcXJ4U1R6dU92S1RzRDU4TXpMR1JEcUl3NHFmNkRxL2R0DQptTVdnRVEvNXFrMFI5b2pDNnRGYUhjVHFxZGYwSk8xcjNPYUxDTDVFMWtYNGl3WTdPSSs2K3k3NVR4ZjROZ0x1DQpLTGcyUWZrTGRLOXorS1JsQm1tWTI4Y0dOZHV5bmtzVFM3QnpyRHQzMjRSMEV0M0V5QzRzQzRlb2JyTGFTbi9ODQpKcU83S3VIaHdheE8wUmFnN0dPNTdZTWFpNy8yV0kxQ09vTlhFOTJUTEg5NXFuMFBJWWVtSlhEVmdsOE8ybzJBDQplcmU2R1JBQXc3UWFjSERpUjdFd1o5UnVwQldNTkFxbHpiNGRFNk13bzZtWkMwcDQrQng0MHZnb2hKbWtiNXhiDQozclJsME0wM0F6RmQwVWpqNGtLOGkzdy9ic0E4cDdMWUlIS3BjQkVjS3MzbGwvV2Y3UTlhVE1LS3NvTDNTdlpYDQpIT0xrS3pYbGc4K1l4blJKZ3ZoS2I4WmFQV3R1V25ZeG5oaUlYVk4yeVBqdjlzcU56NWVmWDJOTDByczhCV0Z4DQp2T0tUcCt5dnp2aVhDeVVpSW9LMFVacHpNTWlaY05LUEdOTHgxWkM4UDJVSTNENWwwL1IvWHo5UGtSTjVveFFzDQppZz09DQotLS0tLUVORCBDRVJUSUZJQ0FURS0tLS0t';
  }

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

  private readonly receiveNotificationMessage = async (notification: Notification) => {
    log('received notification message', notification);
    const emitter: ThreadEventEmitter | undefined = this.emitter;
    if (!notification.encryptedContent) throw new Error('Message did not contain encrypted content');
    if (notification.resource.indexOf('/messages') !== -1) {
      await this.processMessageNotification(notification, emitter);
    } else if (notification.resource.indexOf('/members') !== -1) {
      await this.processMembershipNotification(notification, emitter);
    } else {
      await this.processChatPropertiesNotification(notification, emitter);
    }
    return new signalR.HttpResponse(200);
  };

  private async processMessageNotification(notification: Notification, emitter: ThreadEventEmitter | undefined) {
    const message = await this.decryptMessage<ChatMessage>(notification.encryptedContent);

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

  private async processMembershipNotification(notification: Notification, emitter: ThreadEventEmitter | undefined) {
    const member = await this.decryptMessage<AadUserConversationMember>(notification.encryptedContent);
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

  private async processChatPropertiesNotification(notification: Notification, emitter: ThreadEventEmitter | undefined) {
    const chat = await this.decryptMessage<Chat>(notification.encryptedContent);
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

    // only start timer once. -1 for renewaltimer is semaphore it has stopped.
    if (this.renewalInterval === -1) this.startRenewalTimer();
  };

  private async subscribeToResource(resourcePath: string, changeTypes: ChangeTypes[]) {
    // build subscription request
    const expirationDateTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    ).toISOString();
    const subscriptionDefinition: Subscription = {
      changeType: changeTypes.join(','),
      notificationUrl: 'wss:',
      resource: resourcePath,
      expirationDateTime,
      includeResourceData: true,
      encryptionCertificate: this._publicKey,
      encryptionCertificateId: 'not-a-cert-id',
      clientState: 'wsssecret'
    };

    log('subscribing to changes for ' + resourcePath);
    // send subscription POST to Graph
    const subscription: Subscription = (await Providers.globalProvider.graph.client
      .api(GraphConfig.subscriptionEndpoint)
      .post(subscriptionDefinition)) as Subscription;
    if (!subscription?.notificationUrl) throw new Error('Subscription not created');
    log(subscription);
    subscription.notificationUrl = GraphConfig.adjustNotificationUrl(subscription.notificationUrl);

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
    const renewedSubscription = (await Providers.globalProvider.graph.client
      .api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`)
      .patch({
        expirationDateTime
      })) as Subscription;
    return this.cacheSubscription(renewedSubscription);
  };

  public async createSignalRConnection(notificationUrl: string) {
    const connectionOptions: signalR.IHttpConnectionOptions = {
      accessTokenFactory: this.getToken,
      withCredentials: false
    };
    const connection = new signalR.HubConnectionBuilder()
      .withUrl(notificationUrl, connectionOptions)
      .withAutomaticReconnect()
      .configureLogging(signalR.LogLevel.Information)
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
      tasks.push(Providers.globalProvider.graph.client.api(`${GraphConfig.subscriptionEndpoint}/${s.id}`).delete());
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
