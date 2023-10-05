/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable no-console */
import { Providers } from '@microsoft/mgt-element';
import { HubConnection, HubConnectionBuilder, HubConnectionState, LogLevel } from '@microsoft/signalr';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { ChatMessage, Chat, ConversationMember, AadUserConversationMember } from '@microsoft/microsoft-graph-types';

export const appSettings = {
  functionHost: '', // TODO: improve how this is loaded
  defaultSubscriptionLifetimeInMinutes: 5,
  renewalThreshold: 45, // The number of seconds before subscription expires it will be renewed
  timerInterval: 10, // The number of seconds the timer will check on expiration
  appId: ''
};

const SubscriptionMethods = {
  Create: 'CreateSubscription',
  Renew: 'RenewSubscription'
};

const Return = {
  NewMessage: 'newMessage',
  SubscriptionCreated: 'SubscriptionCreated',
  SubscriptionRenewed: 'SubscriptionRenewed',
  SubscriptionRenewalFailed: 'SubscriptionRenewalFailed',
  SubscriptionCreationFailed: 'SubscriptionCreationFailed',
  SubscriptionRenewalIgnored: 'SubscriptionRenewalIgnored'
};

type ChangeTypes = 'created' | 'updated' | 'deleted';

interface SubscriptionDefinition {
  resource: string;
  expirationTime: Date;
  changeTypes: ChangeTypes[];
  resourceData: boolean;
  signalRConnectionId: string | null;
}

interface Notification {
  subscriptionId: string;
  changeType: ChangeTypes;
  resource: string;
  resourceData: {
    id: string;
    '@odata.type': string;
    '@odata.id': string;
  };
  encryptedContent: string;
}

interface Subscription {
  subscriptionId: string;
  expirationTime: string; // ISO 8601
  resource: string;
}

interface SubscriptionRecord {
  SubscriptionId: string;
  ExpirationTime: string; // ISO 8601
  Resource: string;
}

const loadCachedSubscriptions = (): Subscription[] =>
  JSON.parse(sessionStorage.getItem('graph-subscriptions') || '[]') as Subscription[];

export class GraphNotificationClient {
  private connection?: HubConnection = undefined;
  private renewalInterval = -1;
  private renewalCount = 0;
  private currentUserId = '';

  private subscriptionEmitter: Record<string, ThreadEventEmitter | undefined> = {};
  private tempThreadSubscriptionEmitterMap: Record<string, ThreadEventEmitter> = {};

  private async getToken() {
    const token = await Providers.globalProvider.getAccessToken({
      scopes: [`${appSettings.appId}/.default`]
    });
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  }

  private readonly onRenewalIgnored = (subscriptionRecord: SubscriptionRecord) => {
    console.log('subscription renewalIgnored for subscription ' + subscriptionRecord.SubscriptionId);
  };

  private readonly onRenewalFailed = async (subscriptionId: string) => {
    console.log(`Renewal of subscription ${subscriptionId} failed.`);
    // Something failed to renew the subscription. Create a new one.
    await this.recreateSubscription(subscriptionId);
  };

  private readonly onSubscribeFailed = (subscriptionDefinition: SubscriptionDefinition) => {
    // Something failed when creation the subscription.
    console.log(`Creation of subscription for resource ${subscriptionDefinition.resource} failed.`);
  };

  private buildConnection(): HubConnection {
    return new HubConnectionBuilder()
      .withUrl(appSettings.functionHost + '/api', {
        accessTokenFactory: async () => await this.getToken()
      })
      .withAutomaticReconnect()
      .configureLogging(LogLevel.Information)
      .build();
  }

  private readonly onReconnect = (connectionId: string | undefined) => {
    console.log(`Reconnected. ConnectionId: ${connectionId || 'undefined'}`);
    void this.renewChatSubscriptions();
  };

  private readonly receiveNotificationMessage = async (notification: Notification) => {
    console.log('received notification message', notification);
    const emitter: ThreadEventEmitter | undefined = this.subscriptionEmitter[notification.subscriptionId];
    if (!notification.encryptedContent) throw new Error('Message did not contain encrypted content');
    if (notification.resource.indexOf('/messages') !== -1) {
      await this.processMessageNotification(notification, emitter);
    } else if (notification.resource.indexOf('/members') !== -1) {
      await this.processMembershipNotification(notification, emitter);
    } else {
      await this.processChatPropertiesNotification(notification, emitter);
    }
  };

  private async processMessageNotification(notification: Notification, emitter: ThreadEventEmitter | undefined) {
    const message = await this.decryptNotification<ChatMessage>(notification);
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
    const member = await this.decryptNotification<AadUserConversationMember>(notification);
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
    const chat = await this.decryptNotification<Chat>(notification);
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

  private readonly onSubscribed = (subscriptionRecord: SubscriptionRecord) => {
    console.log(`Subscription created. SubscriptionId: ${subscriptionRecord.SubscriptionId}`);
    this.cacheSubscription(subscriptionRecord);
    this.subscriptionEmitter[subscriptionRecord.SubscriptionId]?.notificationsSubscribedForResource(
      subscriptionRecord.Resource
    );
  };

  private readonly onRenewed = (subscriptionRecord: SubscriptionRecord) => {
    console.log(`Subscription renewed. SubscriptionId: ${subscriptionRecord.SubscriptionId}`);
    this.cacheSubscription(subscriptionRecord);
  };

  private readonly cacheSubscription = (subscriptionRecord: SubscriptionRecord) => {
    console.log(subscriptionRecord);

    // move the event emitter from the temp map to the subscription map
    this.remapEmitter(subscriptionRecord);

    const subscriptions = loadCachedSubscriptions();
    const tempSubscriptions: Subscription[] = subscriptions
      ? subscriptions.filter(
          subscription =>
            subscription.subscriptionId !== subscriptionRecord.SubscriptionId &&
            subscription.resource !== subscriptionRecord.Resource
        )
      : [];

    tempSubscriptions.push({
      subscriptionId: subscriptionRecord.SubscriptionId,
      expirationTime: subscriptionRecord.ExpirationTime,
      resource: subscriptionRecord.Resource
    });

    sessionStorage.setItem('graph-subscriptions', JSON.stringify(tempSubscriptions));

    // only start timer once. -1 for renewaltimer is semaphore it has stopped.
    if (this.renewalInterval === -1) this.startRenewalTimer();
  };

  private remapEmitter(subscriptionRecord: SubscriptionRecord) {
    if (this.tempThreadSubscriptionEmitterMap[subscriptionRecord.Resource]) {
      this.subscriptionEmitter[subscriptionRecord.SubscriptionId] =
        this.tempThreadSubscriptionEmitterMap[subscriptionRecord.Resource];
      delete this.tempThreadSubscriptionEmitterMap[subscriptionRecord.Resource];
    }
  }

  public async closeSignalRConnection() {
    if (this.connection && this.connection?.state === HubConnectionState.Connected) {
      await this.connection?.stop();
    }
  }

  public async reConnectSignalR() {
    if (!this.connection) await this.createSignalConnection();
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    if (this.connection?.state === HubConnectionState.Disconnected) {
      await this.connection?.start();
    }
  }

  private async subscribeToResource(
    resourcePath: string,
    eventEmitter: ThreadEventEmitter,
    changeTypes: ChangeTypes[] = ['created', 'updated', 'deleted']
  ) {
    if (!this.connection) throw new Error('subscribeToResource: SignalR connection is not initialized');

    const token = await this.getToken();

    console.log('subscribing to changes for ' + resourcePath);

    const expirationTime: Date = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );

    const subscriptionDefinition: SubscriptionDefinition = {
      resource: resourcePath,
      expirationTime,
      changeTypes,
      resourceData: true,
      signalRConnectionId: this.connection.connectionId
    };

    // put the eventEmitter into a temporary map so that we can retrieve it when the subscription is created
    this.tempThreadSubscriptionEmitterMap[resourcePath] = eventEmitter;

    await this.connection.send(SubscriptionMethods.Create, subscriptionDefinition, token);

    console.log('Invoked CreateSubscription');
  }

  private readonly startRenewalTimer = () => {
    if (this.renewalInterval !== -1) clearInterval(this.renewalInterval);
    this.renewalInterval = window.setInterval(() => this.renewalTimer(), appSettings.timerInterval * 1000);
    console.log(`Start renewal timer . Id: ${this.renewalInterval}`);
  };

  private readonly renewalTimer = () => {
    const subscriptions = loadCachedSubscriptions();
    if (subscriptions.length === 0) {
      console.log(`No subscriptions found in session state. Stop renewal timer ${this.renewalInterval}.`);
      clearInterval(this.renewalInterval);
      return;
    }

    for (const subscription of subscriptions) {
      const expirationTime = new Date(subscription.expirationTime);
      const now = new Date();
      const diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

      if (diff <= appSettings.renewalThreshold) {
        this.renewalCount++;
        console.log(`Renewing Graph subscription. RenewalCount: ${this.renewalCount}`);
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
    if (!this.connection) {
      throw new Error('No connection');
    }
    clearInterval(this.renewalInterval);
    const token = await this.getToken();

    const expirationTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );

    const subscriptionCache = loadCachedSubscriptions();
    const awaits: Promise<void>[] = [];
    for (const subscription of subscriptionCache) {
      awaits.push(this.connection?.send(SubscriptionMethods.Renew, subscription.subscriptionId, expirationTime, token));
      console.log(`Invoked RenewSubscription ${subscription.subscriptionId}`);
    }
    await Promise.all(awaits);
  };

  private readonly recreateSubscription = async (subscriptionId: string): Promise<void> => {
    console.log('Remove Subscription from session storage.');
    const subscriptionCache = loadCachedSubscriptions();
    if (subscriptionCache?.length > 0) {
      const subscriptions = subscriptionCache.filter(s => s.subscriptionId !== subscriptionId);

      sessionStorage.setItem('graph-subscriptions', JSON.stringify(subscriptions));

      const subscription = subscriptionCache.find(s => s.subscriptionId === subscriptionId);
      const eventEmitter = this.subscriptionEmitter[subscriptionId];
      if (subscription && eventEmitter) {
        await this.subscribeToResource(subscription.resource, eventEmitter);
      }
    }
  };

  private readonly decryptNotification = async <T extends ChatMessage | Chat | ConversationMember[]>(
    notification: Notification
  ): Promise<T> => {
    const token = await this.getToken();

    const response = await fetch(appSettings.functionHost + '/api/GetChatMessageFromNotification', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`
      },
      body: JSON.stringify(notification.encryptedContent)
    });

    if (!response.ok) {
      throw new Error(response.statusText);
    }
    return (await response.json()) as T;
  };

  public async createSignalConnection() {
    const connection = this.buildConnection();

    connection.onreconnected(this.onReconnect);

    connection.on(Return.NewMessage, this.receiveNotificationMessage);

    connection.on('EchoMessage', console.log);

    connection.on(Return.SubscriptionCreated, this.onSubscribed);

    connection.on(Return.SubscriptionRenewed, this.onRenewed);

    connection.on(Return.SubscriptionRenewalIgnored, this.onRenewalIgnored);

    connection.on(Return.SubscriptionRenewalFailed, this.onRenewalFailed);

    connection.on(Return.SubscriptionCreationFailed, this.onSubscribeFailed);

    this.connection = connection;

    await connection.start();
    console.log(connection);
  }

  public async subscribeToChatNotifications(
    userId: string,
    threadId: string,
    eventEmitter: ThreadEventEmitter,
    afterConnection: () => void
  ) {
    if (!this.connection) await this.createSignalConnection();
    afterConnection();
    this.currentUserId = userId;
    const promises: Promise<void>[] = [];
    promises.push(this.subscribeToResource(`/chats/${threadId}/messages`, eventEmitter));
    promises.push(this.subscribeToResource(`/chats/${threadId}/members`, eventEmitter, ['created', 'deleted']));
    promises.push(this.subscribeToResource(`/chats/${threadId}`, eventEmitter, ['updated', 'deleted']));
    await Promise.all(promises);
  }
}
