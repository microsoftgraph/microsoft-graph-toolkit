import {
  ChatMessage as AcsChatMessage,
  ContentSystemMessage,
  MessageThreadProps,
  ErrorBarProps,
  Message,
  SystemMessage
} from '@azure/communication-react';
import { IDynamicPerson, getUserWithPhoto } from '@microsoft/mgt-components';
import {
  ActiveAccountChanged,
  IGraph,
  LoginChangedEvent,
  ProviderState,
  Providers,
  warn
} from '@microsoft/mgt-element';
import { GraphError } from '@microsoft/microsoft-graph-client';
import {
  ChatMessage,
  ChatMessageAttachment,
  ChatRenamedEventMessageDetail,
  MembersAddedEventMessageDetail,
  MembersDeletedEventMessageDetail
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
import { loadChatImage } from './graph.chat';
import { updateMessageContentWithImage } from '../utils/updateMessageContentWithImage';
import { isChatMessage } from '../utils/types';
import { rewriteEmojiContent } from '../utils/rewriteEmojiContent';

// 1x1 grey pixel
const placeholderImageContent =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAMSURBVBhXY7h58yYABRoCjB9qX5UAAAAASUVORK5CYII=';

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

type ChatMessageEvents = MembersAddedEventDetail | MembersRemovedEventDetail | ChatRenamedEventDetail;

const isChatMemberChangeEvent = (
  eventDetail: unknown
): eventDetail is MembersDeletedEventMessageDetail | MembersAddedEventMessageDetail => {
  return (
    typeof (eventDetail as MembersDeletedEventMessageDetail | MembersAddedEventMessageDetail)?.members !== 'undefined'
  );
};

// defines the type of the state object returned from the StatefulGraphChatListClient
export type GraphChatListClient = Pick<MessageThreadProps, 'userId' | 'messages'> & {
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

interface CreatedOn {
  createdOn: Date;
}

/**
 * Simple object comparator function for sorting by createdOn date
 *
 * @param {CreatedOn} a
 * @param {CreatedOn} b
 */
const MessageCreatedComparator = (a: CreatedOn, b: CreatedOn) => a.createdOn.getTime() - b.createdOn.getTime();

type MessageEventType =
  | '#microsoft.graph.membersAddedEventMessageDetail'
  | '#microsoft.graph.membersDeletedEventMessageDetail'
  | '#microsoft.graph.chatRenamedEventMessageDetail';

/**
 * Extended Message type with additional properties.
 */
export type GraphChatMessage = Message & {
  hasUnsupportedContent: boolean;
  rawChatUrl: string;
};
/**
 * Holder type account for async conversion of messages.
 * Some messages need to be written to the UI immediately and receive an async update.
 * Some messages do not have a current value and will be added after the future value is resolved.
 * Some messages do not have a future value and will be added immediately.
 */
interface MessageConversion {
  currentValue?: GraphChatMessage;
  futureValue?: Promise<GraphChatMessage>;
}

export interface ChatListEvent {
  type:
    | 'chatMessageReceived'
    | 'chatMessageDeleted'
    | 'chatMessageEdited'
    | 'chatRenamed'
    | 'memberAdded'
    | 'memberRemoved'
    | 'systemEvent';
  message: ChatMessage;
}

/**
 * Regex to detect and replace image urls using graph requests to supply the image content
 */
const graphImageUrlRegex = /(<img[^>]+)src=(["']https:\/\/graph\.microsoft\.com[^"']*["'])/;

class StatefulGraphChatListClient implements StatefulClient<GraphChatListClient> {
  private readonly _notificationClient: GraphNotificationUserClient;
  private readonly _eventEmitter: ThreadEventEmitter;
  // private readonly _cache: MessageCache;
  private _stateSubscribers: ((state: GraphChatListClient) => void)[] = [];
  private _messageSubscribers: ((messageEvent: ChatListEvent) => void)[] = [];

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    // this._cache = new MessageCache();
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
   * Register a callback to receive chat message updates
   *
   * @param {(messageEvent: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public onChatListEvent(handler: (messageEvent: ChatListEvent) => void): void {
    if (!this._messageSubscribers.includes(handler)) {
      this._messageSubscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving chat message updates
   *
   * @param {(messageEvent: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public offChatListEvent(handler: (messageEvent: ChatListEvent) => void): void {
    const index = this._messageSubscribers.indexOf(handler);
    if (index !== -1) {
      this._messageSubscribers = this._messageSubscribers.splice(index, 1);
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

  private readonly _initialState: GraphChatListClient = {
    status: 'initial',
    activeErrorMessages: [],
    messages: [],
    userId: ''
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
   * Calls each subscriber with the next state to be emitted
   */
  private notifyChatMessageEventChange(message: ChatListEvent) {
    this._messageSubscribers.forEach(handler => handler(message));
  }

  /**
   * Utility to convert a Graph Chat Message to an ACS Chat Message
   * Some graph chat messages can synchronously be converted to ACS chat messages
   * Others require an async call to get the ACS chat message
   * Some messages can provide an initial state for immediate rendering while the async call is in progress
   * This method returns an object containing the current value and a Promise of a future value
   *
   * @private
   * @param {ChatMessage} message
   * @return {*}  {MessageConversion}
   * @memberof StatefulGraphChatListClient
   */
  private convertChatMessage(message: ChatMessage): MessageConversion {
    switch (message.messageType) {
      case 'message':
        return this.graphChatMessageToAcsChatMessage(message, this.userId);
      case 'systemEventMessage':
      case 'unknownFutureValue':
        return { futureValue: this.buildSystemContentMessage(message) } as MessageConversion;
      default:
        throw new Error(`Unknown message type ${message.messageType?.toString() || 'undefined'}`);
    }
  }

  /**
   * Convert an event type to the name of an icon registered in registerIcons.tsx
   *
   * @param eventType
   * @returns
   */
  private readonly resolveIcon = (eventType: MessageEventType): string => {
    switch (eventType) {
      case '#microsoft.graph.membersAddedEventMessageDetail':
        return 'add-friend';
      case '#microsoft.graph.membersDeletedEventMessageDetail':
        return 'left-chat';
      case '#microsoft.graph.chatRenamedEventMessageDetail':
        return 'edit-svg';
      default:
        return 'Unknown';
    }
  };

  private async buildSystemContentMessage(message: ChatMessage): Promise<SystemMessage> {
    const eventDetail = message.eventDetail as ChatMessageEvents;
    let messageContent = '';
    const awaits: Promise<IDynamicPerson>[] = [];
    const initiatorId = eventDetail.initiator?.user?.id;
    // we're using getUserWithPhoto here because we want to tap into the caching that mgt-person uses to cut down on graph calls
    if (initiatorId) {
      awaits.push(getUserWithPhoto(this.graph, initiatorId));
      const userIds: string[] = [];
      if (isChatMemberChangeEvent(eventDetail)) {
        eventDetail.members?.reduce((acc, m) => {
          if (typeof m.id === 'string') {
            acc.push(m.id);
          }
          return acc;
        }, userIds);
      }
      for (const id of userIds ?? []) {
        awaits.push(getUserWithPhoto(this.graph, id));
      }
      const people = await Promise.all(awaits);
      const userNames = userIds
        ?.filter(id => id !== initiatorId)
        .map(m => this.getUserName(m, people))
        .join(', ');
      const initiatorUsername = this.getUserName(initiatorId, people);
      switch (eventDetail['@odata.type']) {
        case '#microsoft.graph.membersAddedEventMessageDetail':
          messageContent = `${initiatorUsername} added ${userNames}`;
          break;
        case '#microsoft.graph.membersDeletedEventMessageDetail':
          messageContent = `${initiatorUsername} removed ${userNames}`;
          break;
        case '#microsoft.graph.chatRenamedEventMessageDetail':
          messageContent = eventDetail.chatDisplayName
            ? `${initiatorUsername} renamed the chat to ${eventDetail.chatDisplayName}`
            : `${initiatorUsername} removed the group name for this conversation`;
          break;
        // TODO: move this default case to a console.warn before release and emit an empty message
        // it's here to help us catch messages we have't handled yet
        default:
          // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
          warn(`Unknown system message type ${eventDetail['@odata.type']}

detail: ${JSON.stringify(eventDetail)}`);
      }
    }

    const result: ContentSystemMessage = {
      createdOn: message.createdDateTime ? new Date(message.createdDateTime) : new Date(),
      messageId: message.id || '',
      systemMessageType: 'content',
      messageType: 'system',
      iconName: this.resolveIcon(eventDetail['@odata.type']),
      content: messageContent
    };

    return result;
  }

  private getUserName(userId: string, people: IDynamicPerson[]) {
    return people.find(p => p.id === userId)?.displayName || 'unknown user';
  }

  /*
   * Helper method to set the content of a message to show deletion
   */
  private readonly setDeletedContent = (message: AcsChatMessage) => {
    message.content = '<em>This message has been deleted.</em>';
    message.contentType = 'html';
  };

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
        default:
          return 'systemEvent';
      }
    }

    return 'chatMessageReceived';
  }

  /*
   * Event handler to be called when a new message is received by the notification service
   */
  private readonly onMessageReceived = async (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: this.getSystemMessageType(message) });

      // await this._cache.cacheMessage(message.chatId, message);
      const messageConversion = this.convertChatMessage(message);
      const acsMessage = messageConversion.currentValue;
      this.updateMessages(acsMessage);
      if (messageConversion.futureValue) {
        // if we have a future value then we need to wait for it to resolve before we can send the read receipt
        const futureMessageState = await messageConversion.futureValue;
        this.updateMessages(futureMessageState);
      }
    }
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private readonly onMessageDeleted = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageDeleted' });

      // void this._cache.deleteMessage(message.chatId, message);
      this.notifyStateChange((draft: GraphChatListClient) => {
        const draftMessage = draft.messages.find(m => m.messageId === message.id) as AcsChatMessage;
        // TODO: confirm if we should show the deleted content message in all cases or only when the message was deleted by the current user
        if (draftMessage) this.setDeletedContent(draftMessage);
      });
    }
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private readonly onMessageEdited = async (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageEdited' });

      // await this._cache.cacheMessage(message.chatId, message);
      const messageConversion = this.convertChatMessage(message);
      this.updateMessages(messageConversion.currentValue);
      if (messageConversion.futureValue) {
        const eventualState = await messageConversion.futureValue;
        this.updateMessages(eventualState);
      }
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
      draft.messages = [];
      draft.status = 'initial'; // no message?
    });
  }

  private updateUserInfo() {
    this.updateCurrentUserId();
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
   * @memberof StatefulGraphChatListClient
   */
  private async updateUserSubscription() {
    // avoid subscribing to a resource with an empty chatId
    if (this.userId && this._sessionId) {
      // reset state to initial
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'initial';
        draft.messages = [];
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

  /**
   * Update the state with given message either replacing an existing message matching on the id or adding to the list
   *
   * @private
   * @param {(GraphChatMessage)} [message]
   * @return {*}
   * @memberof StatefulGraphChatListClient
   */
  private updateMessages(message?: GraphChatMessage) {
    if (!message) return;
    this.notifyStateChange((draft: GraphChatListClient) => {
      const index = draft.messages.findIndex(m => m.messageId === message.messageId);
      // this message is not already in thread so just add it
      if (index === -1) {
        // sort to ensure that messages are in the correct order should we get messages out of order
        draft.messages = draft.messages.concat(message).sort(MessageCreatedComparator);
      } else {
        // replace the existing version of the message with the new one
        draft.messages.splice(index, 1, message);
      }
    });
  }

  private graphImageMatch(messageContent: string): RegExpMatchArray | null {
    return messageContent.match(graphImageUrlRegex);
  }

  private processMessageContent(graphMessage: ChatMessage, currentUser: string): MessageConversion {
    const conversion: MessageConversion = {};
    // using a record here lets us track which image in the content each request is for
    const futureImages: Record<number, Promise<string | null>> = {};
    let messageResult = graphMessage.body?.content ?? '';
    const messageId = graphMessage.id ?? '';
    let index = 0;
    let match = this.graphImageMatch(messageResult);
    while (match) {
      // note that the regex to replace the placeholder requires that the id be before and adjacent to the src attribute
      messageResult = messageResult.replace(
        graphImageUrlRegex,
        `$1id="${messageId}-${index}" src="${placeholderImageContent}"`
      );
      const graphImageUrl = match[2].replace(/["']/g, '');
      // collect promises for the image requests so that we can update the message content once they are all resolved
      futureImages[index] = loadChatImage(this.graph, graphImageUrl);
      index++;
      match = this.graphImageMatch(messageResult);
    }
    let placeholderMessage = this.buildAcsMessage(graphMessage, currentUser, messageId, messageResult);
    conversion.currentValue = placeholderMessage;
    // local function to update the message with data from each of the resolved image requests
    const updateMessage = async () => {
      await Promise.all(Object.values(futureImages));
      for (const [imageIndex, futureImage] of Object.entries(futureImages)) {
        const image = await futureImage;
        if (image && isChatMessage(placeholderMessage)) {
          placeholderMessage = {
            ...placeholderMessage,
            ...{
              content: updateMessageContentWithImage(placeholderMessage.content ?? '', imageIndex, messageId, image)
            }
          };
        }
      }
      return placeholderMessage;
    };
    conversion.futureValue = updateMessage();
    return conversion;
  }

  private graphChatMessageToAcsChatMessage(graphMessage: ChatMessage, currentUser: string): MessageConversion {
    const messageId = graphMessage?.id ?? '';
    if (!messageId) {
      throw new Error('Cannot convert graph message to ACS message. No ID found on graph message');
    }
    let content = graphMessage.body?.content ?? 'undefined';
    let result: MessageConversion = {};
    // do simple emoji replacement first
    content = rewriteEmojiContent(content);

    const imageMatch = this.graphImageMatch(content ?? '');
    if (imageMatch) {
      // if the message contains an image, we need to fetch the image and replace the placeholder
      result = this.processMessageContent(graphMessage, currentUser);
    } else {
      result.currentValue = this.buildAcsMessage(graphMessage, currentUser, messageId, content);
    }
    return result;
  }

  private hasUnsupportedContent(content: string, attachments: ChatMessageAttachment[]): boolean {
    const unsupportedContentTypes = [
      'application/vnd.microsoft.card.codesnippet',
      'application/vnd.microsoft.card.fluid',
      'application/vnd.microsoft.card.fluidEmbedCard',
      'reference'
    ];
    const isUnsupported: boolean[] = [];

    if (attachments.length) {
      for (const attachment of attachments) {
        const contentType = attachment?.contentType ?? '';
        isUnsupported.push(unsupportedContentTypes.includes(contentType));
      }
    } else {
      // checking content with <attachment> tags
      const unsupportedContentRegex = /<\/?attachment>/gim;
      const contentUnsupported = Boolean(content) && unsupportedContentRegex.test(content);
      isUnsupported.push(contentUnsupported);
    }
    return isUnsupported.every(e => e === true);
  }

  private buildAcsMessage(
    graphMessage: ChatMessage,
    currentUser: string,
    messageId: string,
    content: string
  ): GraphChatMessage {
    const senderId = graphMessage.from?.user?.id || undefined;
    const chatId = graphMessage?.chatId ?? '';
    const id = graphMessage?.id ?? '';
    const chatUrl = `https://teams.microsoft.com/l/message/${chatId}/${id}?context={"contextType":"chat"}`;
    const attachments = graphMessage?.attachments ?? [];

    let messageData: GraphChatMessage = {
      messageId,
      contentType: graphMessage.body?.contentType ?? 'text',
      messageType: 'chat',
      content,
      senderDisplayName: graphMessage.from?.user?.displayName ?? undefined,
      createdOn: new Date(graphMessage.createdDateTime ?? Date.now()),
      editedOn: graphMessage.lastEditedDateTime ? new Date(graphMessage.lastEditedDateTime) : undefined,
      senderId,
      mine: senderId === currentUser,
      status: 'seen',
      attached: 'top',
      hasUnsupportedContent: this.hasUnsupportedContent(content, attachments),
      rawChatUrl: chatUrl
    };
    if (graphMessage?.policyViolation) {
      messageData = Object.assign(messageData, {
        messageType: 'blocked',
        link: 'https://go.microsoft.com/fwlink/?LinkId=2132837'
      });
    }
    return messageData;
  }

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', (message: ChatMessage) => void this.onMessageReceived(message));
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', (message: ChatMessage) => void this.onMessageEdited(message));
  }

  /**
   * Provided the graph instance for the component with the correct SDK version decoration
   *
   * @readonly
   * @private
   * @type {IGraph}
   * @memberof StatefulGraphChatListClient
   */
  private get graph(): IGraph {
    return graph('mgt-chat');
  }
}

export { StatefulGraphChatListClient };
