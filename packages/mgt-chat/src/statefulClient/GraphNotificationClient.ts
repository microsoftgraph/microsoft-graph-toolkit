import { Providers } from '@microsoft/mgt-element';
import * as signalR from '@microsoft/signalr';
import { graphChatMessageToACSChatMessage } from './acs.chat';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { ChatMessage } from '@microsoft/microsoft-graph-types';

const appSettings = {
  functionHost: 'https://mgtgnbfunc.azurewebsites.net', // TODO: improve how this is loaded
  defaultSubscriptionLifetimeInMinutes: 5,
  renewalThreshold: 45, //The number of seconds before subscription expires it will be renewed
  timerInterval: 10 //The number of seconds the timer will check on expiration
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

type SubscriptionDefinition = {
  resource: string;
  expirationTime: Date;
  changeTypes: ChangeTypes[];
  resourceData: boolean;
  signalRConnectionId: string | null;
};

type Notification = {
  subscriptionId: string;
  changeType: ChangeTypes;
  encryptedContent: string;
};

type Subscription = {
  subscriptionId: string;
  expirationTime: string; // ISO 8601
  resource: string;
};

type SubscriptionRecord = {
  SubscriptionId: string;
  ExpirationTime: string; // ISO 8601
  Resource: string;
};

const loadCachedSubscriptions = (): Subscription[] => JSON.parse(sessionStorage.getItem('graph-subscriptions') || '[]');

export class GraphNotificationClient {
  private connection?: signalR.HubConnection = undefined;
  private renewalInterval = -1;
  private renewalCount = 0;
  private currentUserId = '';

  private subscriptionEmitter: Record<string, ThreadEventEmitter | undefined> = {};
  private tempThreadSubscriptionEmitterMap: Record<string, ThreadEventEmitter> = {};

  private async getToken() {
    const token = await Providers.globalProvider.getAccessToken({
      scopes: ['de25c6a1-c9e7-4681-9288-eba1b93446fb/.default']
    });
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  }

  private onRenewalIgnored = (subscriptionRecord: SubscriptionRecord) => {
    console.log('subscription renewalIgnored for subscription ' + subscriptionRecord.SubscriptionId);
  };

  private onRenewalFailed = (subscriptionId: string) => {
    console.log(`Renewal of subscription ${subscriptionId} failed.`);
    //Something failed to renew the subscription. Create a new one.
    this.recreateSubscription(subscriptionId);
  };

  private onSubscribeFailed = (subscriptionDefinition: SubscriptionDefinition) => {
    //Something failed when creation the subscription.
    console.log(`Creation of subscription for resource ${subscriptionDefinition.resource} failed.`);
  };

  private buildConnection() {
    return new signalR.HubConnectionBuilder()
      .withUrl(appSettings.functionHost + '/api', {
        accessTokenFactory: async () => await this.getToken()
      })
      .withAutomaticReconnect()
      .configureLogging(signalR.LogLevel.Information)
      .build();
  }

  private onReconnect = (connectionId: string | undefined) => {
    console.log(`Reconnected. ConnectionId: ${connectionId}`);
    this.renewChatSubscriptions();
  };

  private receiveNotificationMessage = async (notification: Notification) => {
    console.log('received notification message', notification);
    if (!notification.encryptedContent) throw new Error('Message did not contain encrypted content');

    const decryptedContent = await this.decryptChatMessageFromNotification(notification);
    const message = graphChatMessageToACSChatMessage(decryptedContent, this.currentUserId);

    const emitter: ThreadEventEmitter | undefined = this.subscriptionEmitter[notification.subscriptionId];
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
  };

  private onSubscribed = async (subscriptionRecord: SubscriptionRecord) => {
    console.log(`Subscription created. SubscriptionId: ${subscriptionRecord.SubscriptionId}`);
    this.cacheSubscription(subscriptionRecord);
    this.subscriptionEmitter[subscriptionRecord.SubscriptionId]?.chatMessageNotificationsSubscribed(
      subscriptionRecord.Resource
    );
  };

  private onRenewed = async (subscriptionRecord: SubscriptionRecord) => {
    console.log(`Subscription renewed. SubscriptionId: ${subscriptionRecord.SubscriptionId}`);
    this.cacheSubscription(subscriptionRecord);
  };

  private cacheSubscription = (subscriptionRecord: SubscriptionRecord) => {
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

    //only start timer once. -1 for renewaltimer is semaphore it has stopped.
    if (this.renewalInterval === -1) this.startRenewalTimer();
  };

  private remapEmitter(subscriptionRecord: SubscriptionRecord) {
    if (this.tempThreadSubscriptionEmitterMap[subscriptionRecord.Resource]) {
      this.subscriptionEmitter[subscriptionRecord.SubscriptionId] = this.tempThreadSubscriptionEmitterMap[
        subscriptionRecord.Resource
      ];
      delete this.tempThreadSubscriptionEmitterMap[subscriptionRecord.Resource];
    }
  }

  private async subscribeToResource(resourcePath: string, eventEmitter: ThreadEventEmitter) {
    if (!this.connection) throw new Error('SignalR connection not initialized');

    const token = await this.getToken();

    console.log('subscribing to changes for ' + resourcePath);

    const expirationTime: Date = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );

    const subscriptionDefinition: SubscriptionDefinition = {
      resource: resourcePath,
      expirationTime: expirationTime,
      changeTypes: ['created', 'updated', 'deleted'],
      resourceData: true,
      signalRConnectionId: this.connection.connectionId
    };

    // put the eventEmitter into a temporary map so that we can retrieve it when the subscription is created
    this.tempThreadSubscriptionEmitterMap[resourcePath] = eventEmitter;

    await this.connection.send(SubscriptionMethods.Create, subscriptionDefinition, token);

    console.log('Invoked CreateSubscription');
  }

  private startRenewalTimer = () => {
    if (this.renewalInterval !== -1) clearInterval(this.renewalInterval);
    this.renewalInterval = window.setInterval(() => this.renewalTimer(), appSettings.timerInterval * 1000);
    console.log(`Start renewal timer . Id: ${this.renewalInterval}`);
  };

  private renewalTimer = () => {
    const subscriptions = loadCachedSubscriptions();
    if (subscriptions.length === 0) {
      console.log(`No subscriptions found in session state. Stop renewal timer ${this.renewalInterval}.`);
      clearInterval(this.renewalInterval);
      return;
    }

    for (const subscription of subscriptions) {
      const expirationTime = new Date(subscription.expirationTime);
      const now = new Date();
      var diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

      if (diff <= appSettings.renewalThreshold) {
        this.renewalCount++;
        console.log(`Renewing Graph subscription. RenewalCount: ${this.renewalCount}`);
        // stop interval to prevent new invokes until refresh is ready.
        clearInterval(this.renewalInterval);
        this.renewalInterval = -1;
        this.renewChatSubscriptions();
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

    let expirationTime = new Date(new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000);

    const subscriptionCache = loadCachedSubscriptions();
    const awaits: Promise<void>[] = [];
    for (const subscription of subscriptionCache) {
      awaits.push(this.connection?.send(SubscriptionMethods.Renew, subscription.subscriptionId, expirationTime, token));
      console.log(`Invoked RenewSubscription ${subscription.subscriptionId}`);
    }
    await Promise.all(awaits);
  };

  private recreateSubscription = (subscriptionId: string) => {
    console.log('Remove Subscription from session storage.');
    const subscriptionCache = loadCachedSubscriptions();
    if (subscriptionCache?.length > 0) {
      const subscriptions = subscriptionCache.filter(subscription => subscription.subscriptionId !== subscriptionId);

      sessionStorage.setItem('graph-subscriptions', JSON.stringify(subscriptions));

      const subscription = subscriptionCache.find(s => s.subscriptionId === subscriptionId);
      const eventEmitter = this.subscriptionEmitter[subscriptionId];
      if (subscription && eventEmitter) {
        this.subscribeToResource(subscription.resource, eventEmitter);
      }
    }
  };

  private decryptChatMessageFromNotification = async (notification: Notification): Promise<ChatMessage> => {
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
    return await response.json();
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

  public async subscribeToChatNotifications(userId: string, threadId: string, eventEmitter: ThreadEventEmitter) {
    if (!this.connection) await this.createSignalConnection();

    this.currentUserId = userId;
    return this.subscribeToResource(`chats/${threadId}/messages`, eventEmitter);
  }
}
