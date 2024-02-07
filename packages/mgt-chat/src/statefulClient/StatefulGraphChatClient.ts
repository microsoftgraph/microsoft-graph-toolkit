/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  ChatMessage as AcsChatMessage,
  ContentSystemMessage,
  Message,
  MessageThreadProps,
  SendBoxProps,
  SystemMessage
} from '@azure/communication-react';
import { IDynamicPerson, getUserWithPhoto } from '@microsoft/mgt-components';
import {
  ActiveAccountChanged,
  IGraph,
  LoginChangedEvent,
  ProviderState,
  Providers,
  error,
  log,
  warn
} from '@microsoft/mgt-element';
import {
  AadUserConversationMember,
  Chat,
  ChatMessage,
  ChatMessageAttachment,
  ChatRenamedEventMessageDetail,
  MembersAddedEventMessageDetail,
  MembersDeletedEventMessageDetail,
  ChatMessageMention,
  NullableOption
} from '@microsoft/microsoft-graph-types';
import { produce } from 'immer';
import { v4 as uuid } from 'uuid';
import { currentUserId, currentUserName } from '../utils/currentUser';
import { graph } from '../utils/graph';
import { MessageCache } from './Caching/MessageCache';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationClient } from './GraphNotificationClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import {
  MessageCollection,
  addChatMembers,
  deleteChatMessage,
  loadChat,
  loadChatImage,
  loadChatThread,
  loadChatThreadDelta,
  loadMoreChatMessages,
  removeChatMember,
  sendChatMessage,
  updateChatMessage,
  updateChatTopic
} from './graph.chat';
import { updateMessageContentWithImage } from '../utils/updateMessageContentWithImage';
import { GraphChatClientStatus, isChatMessage } from '../utils/types';
import { rewriteEmojiContentToHTML } from '../utils/rewriteEmojiContent';

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

// defines the type of the state object returned from the StatefulGraphChatClient
export type GraphChatClient = Pick<
  MessageThreadProps,
  | 'userId'
  | 'messages'
  | 'participantCount'
  | 'disableEditing'
  | 'onLoadPreviousChatMessages'
  | 'numberOfChatMessagesToReload'
  | 'onUpdateMessage'
  | 'onDeleteMessage'
> &
  Pick<SendBoxProps, 'onSendMessage'> & {
    status: GraphChatClientStatus;
    chat?: Chat;
  } & {
    participants: AadUserConversationMember[];
    onAddChatMembers: (userIds: string[], history?: Date) => Promise<void>;
    onRemoveChatMember: (membershipId: string) => Promise<void>;
    onRenameChat: (topic: string | null) => Promise<void>;
    mentions: NullableOption<ChatMessageMention[]>;
    activeErrorMessages: Error[];
  };

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

const ChatMessageCreatedComparator = (a: ChatMessage, b: ChatMessage): number => {
  if (a.createdDateTime && b.createdDateTime) {
    if (a.createdDateTime === b.createdDateTime) return 0;
    if (a.createdDateTime > b.createdDateTime) return 1;
    return -1;
  } else if (a.createdDateTime) {
    return 1;
  }
  return -1;
};

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

/**
 * Regex to detect and replace image urls using graph requests to supply the image content
 */
const graphImageUrlRegex = /(<img[^>]+)src=(["']https:\/\/graph\.microsoft\.com[^"']*["'])/;

class StatefulGraphChatClient implements StatefulClient<GraphChatClient> {
  private readonly _eventEmitter: ThreadEventEmitter;
  private readonly _cache: MessageCache;
  private _graph: IGraph = Providers.globalProvider.graph;
  private _notificationClient: GraphNotificationClient;
  private _subscribers: ((state: GraphChatClient) => void)[] = [];
  private get _messagesPerCall() {
    return 5;
  }
  private _nextLink?: string;
  private _chat?: Chat = undefined;
  private _userDisplayName = '';

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    this._notificationClient = new GraphNotificationClient(this._eventEmitter, this._graph);
    this._cache = new MessageCache();
  }

  /**
   * Updates the provider graph SDK client and the notification client. By now,
   * an assumption is made that the Providers.globalProvider.graph is available
   * and won't throw a TypeError when you access graph.forComponent.
   */
  private updateGraphNotificationClient() {
    this._graph = graph('mgt-chat', GraphConfig.version);
    this._notificationClient = new GraphNotificationClient(this._eventEmitter, this.graph);
  }

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public async tearDown() {
    await this._notificationClient?.tearDown();
  }

  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public onStateChange(handler: (state: GraphChatClient) => void): void {
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
  public offStateChange(handler: (state: GraphChatClient) => void): void {
    const index = this._subscribers.indexOf(handler);
    if (index !== -1) {
      this._subscribers = this._subscribers.splice(index, 1);
    }
  }

  /**
   * Calls each subscriber with the next state to be emitted
   *
   * @param recipe - a function which produces the next state to be emitted
   */
  private notifyStateChange(recipe: (draft: GraphChatClient) => void) {
    this._state = produce(this._state, recipe);
    this._subscribers.forEach(handler => handler(this._state));
  }

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  public getState(): GraphChatClient {
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
        this.updateGraphNotificationClient();
        // load messages?
        // configure subscriptions
        // emit new state;
        if (this.chatId) {
          void this.updateFollowedChat();
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
    await this.updateFollowedChat();
  };

  private clearCurrentUserMessages() {
    log('clear');
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.messages = [];
      draft.participants = [];
      draft.chat = undefined;
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
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.userId = userId;
    });
  }

  /**
   * Current chat ID.
   */
  private _chatId = '';

  /**
   * Get the current chat ID.
   */
  public get chatId() {
    return this._chatId;
  }

  /**
   * Set the current chat ID and tries to get the chat data.
   */
  private set chatId(value: string) {
    // take no action if the chatId is the same
    if (value && this._chatId === value) {
      return;
    }
    this._chatId = value;
    if (this._chatId && this._sessionId) {
      void this.updateFollowedChat();
    }
  }

  private _sessionId: string | undefined;

  public subscribeToChat(chatId: string, sessionId: string) {
    if (chatId && sessionId) {
      this._sessionId = sessionId;
      this.chatId = chatId;
    }
  }

  public setStatus(status: GraphChatClientStatus) {
    this.notifyStateChange((draft: GraphChatClient) => {
      log(`status = ${status}`);
      draft.status = status;
      draft.chat = { topic: 'Unknown' };
    });
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatClient
   */
  private async updateFollowedChat() {
    // avoid subscribing to a resource with an empty chatId
    if (this.chatId && this._sessionId) {
      // reset state to initial
      log('rset');
      this.notifyStateChange((draft: GraphChatClient) => {
        draft.status = 'initial';
        draft.messages = [];
        draft.mentions = [];
        draft.chat = undefined;
        draft.participants = [];
        draft.activeErrorMessages = [];
      });
      // Subscribe to notifications for messages
      this.notifyStateChange((draft: GraphChatClient) => {
        draft.status = 'creating server connections';
      });
      try {
        // Prefer sequential promise resolving to catch loading message errors
        // TODO: in parallel promise resolving, find out how to trigger different
        // TODO: state for failed subscriptions in GraphChatClient.onSubscribeFailed
        const tasks: Promise<unknown>[] = [this.loadChatData()];
        // subscribing to notifications will trigger the chatMessageNotificationsSubscribed event
        // this client will then load the chat and messages when that event listener is called
        tasks.push(this._notificationClient.subscribeToChatNotifications(this._chatId, this._sessionId));
        await Promise.all(tasks);
      } catch (e) {
        this.updateErrorState(e as Error);
      }
    } else {
      this.notifyStateChange((draft: GraphChatClient) => {
        draft.status = 'no chat id';
        draft.chat = { topic: 'Unknown' };
      });
    }
  }

  private async loadChatData() {
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'loading messages';
    });
    try {
      this._chat = await loadChat(this.graph, this.chatId);
    } catch (e) {
      return Promise.reject(e);
    }
    const cachedMessages = await this._cache.loadMessages(this._chatId);
    if (cachedMessages) {
      // use currently cached data
      await this.writeMessagesToState(cachedMessages);
      // load delta messages
      const deltaMessages = await this.loadDeltaData(this._chatId, cachedMessages.lastModifiedDateTime);
      // add delta messages to cache
      const updatedState = await this._cache.cacheMessages(this._chatId, deltaMessages);
      // writeMessagesToState concats with existing state, need to be careful not to create duplicate messages
      updatedState.value = deltaMessages;
      // update state
      await this.writeMessagesToState(updatedState);
    } else {
      const messages: MessageCollection = await loadChatThread(this.graph, this._chatId, this._messagesPerCall);
      await this._cache.cacheMessages(this._chatId, messages.value, true, messages.nextLink);
      await this.writeMessagesToState(messages);
    }
  }

  private async loadDeltaData(chatId: string, lastModified: string): Promise<ChatMessage[]> {
    const result: ChatMessage[] = [];
    let response = await loadChatThreadDelta(this.graph, chatId, lastModified, this._messagesPerCall);
    result.push(...response.value);
    while (response.nextLink) {
      response = await loadMoreChatMessages(this.graph, response.nextLink);
      result.push(...response.value);
    }
    return result.sort(ChatMessageCreatedComparator);
  }

  /**
   * Handles updating state based on the response to a graph request for a collection of messages
   * Handles messages that can be imedately rendered and those that require async calls to fetch data
   *
   * @private
   * @param {MessageCollection} messages
   * @memberof StatefulGraphChatClient
   */
  private async writeMessagesToState(messages: MessageCollection) {
    this._nextLink = messages.nextLink;

    const messageConversions = messages.value
      // trying to filter out messages on the graph request causes a 400
      // deleted messages are returned as messages with no content, which we can't filter on the graph request
      // so we filter them out here
      .filter(m => m.body?.content)
      // This gives us both current and eventual values for each message
      .map(m => this.convertChatMessage(m));

    // Collect mentions
    const mentions: NullableOption<ChatMessageMention[]> = messages.value
      .map(m => m.mentions)
      .filter(m => m?.length) as NullableOption<ChatMessageMention[]>;

    // update the state with the current values
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.participants = this._chat?.members || [];
      draft.participantCount = draft.participants.length;
      const initialMessages: Message[] = [];
      draft.messages = draft.messages
        .concat(
          messageConversions
            .map(m => m.currentValue)
            // need to use a reduce here to filter out undefined values in a way that TypeScript understands
            .reduce((acc, val) => {
              if (val) acc.push(val);
              return acc;
            }, initialMessages)
        )
        .sort(MessageCreatedComparator);
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
      draft.status = this._nextLink ? 'loading messages' : !draft.messages.length ? 'no messages' : 'ready';
      draft.chat = this._chat;
      // Keep updating if there was a next link.
      draft.mentions = draft.mentions?.concat(
        ...Array.from(mentions as Iterable<ChatMessageMention>)
      ) as NullableOption<ChatMessageMention[]>;
    });
    const futureMessages = messageConversions.filter(m => m.futureValue).map(m => m.futureValue);
    // if there are eventual future values, wait for them to resolve and update the state
    if (futureMessages.length > 0) {
      (await Promise.all(futureMessages)).forEach(m => {
        this.updateMessages(m);
      });
    }
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
   * @memberof StatefulGraphChatClient
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

  /**
   * Async callback to load more messages
   *
   * @returns true if there are no more messages to load
   */
  private readonly loadMoreMessages = async () => {
    if (!this._nextLink) {
      return true;
    }
    const messages: MessageCollection = await loadMoreChatMessages(this.graph, this._nextLink);
    await this._cache.cacheMessages(this._chatId, messages.value, true, messages.nextLink);
    await this.writeMessagesToState(messages);
    // return true when there are no more messages to load
    return !this._nextLink;
  };

  /**
   * Send a message to the chat thread
   *
   * @param {string} content - the content of the message
   * @memberof StatefulGraphChatClient
   */
  public sendMessage = async (content: string) => {
    if (!content) return;

    const pendingId = uuid();

    // add a pending message to the state.
    this.notifyStateChange((draft: GraphChatClient) => {
      const pendingMessage: Message = {
        clientMessageId: pendingId,
        messageId: pendingId,
        contentType: 'text',
        messageType: 'chat',
        content,
        senderDisplayName: this._userDisplayName,
        createdOn: new Date(),
        senderId: this.userId,
        mine: true,
        status: 'sending'
      };
      draft.messages.push(pendingMessage);
    });
    try {
      // send message
      const chat: ChatMessage = await sendChatMessage(this.graph, this.chatId, content);
      // emit new state
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftIndex = draft.messages.findIndex(m => m.messageId === pendingId);
        const message = this.graphChatMessageToAcsChatMessage(chat, this.userId).currentValue;
        // we only need use the current value of the message
        // this message can't have a future value as it's not been sent yet
        if (message) draft.messages.splice(draftIndex, 1, message);
      });
    } catch (e) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftMessage = draft.messages.find(m => m.messageId === pendingId);
        (draftMessage as AcsChatMessage).status = 'failed';
      });
      throw new Error('Failed to send message');
    }
  };

  /*
   * Helper method to set the content of a message to show deletion
   */
  private readonly setDeletedContent = (message: AcsChatMessage) => {
    message.content = '<em>This message has been deleted.</em>';
    message.contentType = 'html';
  };

  /**
   * Handler to delete a message
   *
   * @param messageId id of the message to be deleted, this is the clientMessageId when triggered by the re-send action on a failed message, or the messageId when triggered by the delete action on a sent message
   * @returns {Promise<void>}
   */
  public deleteMessage = async (messageId: string): Promise<void> => {
    if (!messageId) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as AcsChatMessage;
    // only messages not persisted to graph should have a clientMessageId
    const uncommitted = this._state.messages.find(
      m => (m as AcsChatMessage).clientMessageId === messageId
    ) as AcsChatMessage;
    if (message?.mine) {
      try {
        // uncommitted messages are not persisted to the graph, so don't call graph when deleting them
        if (!uncommitted) {
          await deleteChatMessage(this.graph, this.chatId, messageId);
        }
        this.notifyStateChange((draft: GraphChatClient) => {
          const draftMessage = draft.messages.find(m => m.messageId === messageId) as AcsChatMessage;
          if (draftMessage.clientMessageId) {
            // just remove messages that were not saved to the graph
            draft.messages.splice(draft.messages.indexOf(draftMessage), 1);
          } else {
            // show deleted messages which have been persisted to the graph as deleted in the UI
            this.setDeletedContent(draftMessage);
          }
        });
      } catch (e) {
        // TODO: How do we handle failed deletes?
      }
    }
  };

  /**
   * Update a message in the thread
   *
   * @param {string} messageId Id of the message to be updated
   * @param {string} content new content of the message
   * @memberof StatefulGraphChatClient
   */
  public updateMessage = async (messageId: string, content: string) => {
    if (!messageId || !content) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as AcsChatMessage;
    if (message?.mine && message.content) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const updating = draft.messages.find(m => m.messageId === messageId) as AcsChatMessage;
        if (updating) {
          updating.content = content;
          updating.status = 'sending';
        }
      });
      try {
        await updateChatMessage(this.graph, this.chatId, messageId, content);
        this.notifyStateChange((draft: GraphChatClient) => {
          const updated = draft.messages.find(m => m.messageId === messageId) as AcsChatMessage;
          updated.status = 'delivered';
        });
      } catch (e) {
        this.notifyStateChange((draft: GraphChatClient) => {
          const updating = draft.messages.find(m => m.messageId === messageId) as AcsChatMessage;
          updating.status = 'failed';
        });
        throw new Error('Failed to update message');
      }
    }
  };

  /*
   * Event handler to be called when a new message is received by the notification service
   */
  private readonly onMessageReceived = async (message: ChatMessage) => {
    this.updateMentions(message);
    await this._cache.cacheMessage(this._chatId, message);
    const messageConversion = this.convertChatMessage(message);
    const acsMessage = messageConversion.currentValue;
    this.updateMessages(acsMessage);
    if (messageConversion.futureValue) {
      // if we have a future value then we need to wait for it to resolve before we can send the read receipt
      const futureMessageState = await messageConversion.futureValue;
      this.updateMessages(futureMessageState);
    }
  };

  /**
   * When you receive a new message, check if there are any mentions and update
   * the state. This will allow to match users to mentions them during rendering.
   *
   * @param newMessage from teams.
   * @returns
   */
  private readonly updateMentions = (newMessage: ChatMessage) => {
    if (!newMessage) return;
    this.notifyStateChange((draft: GraphChatClient) => {
      const mentions = newMessage?.mentions ?? [];
      draft.mentions = draft.mentions?.concat(
        ...Array.from(mentions as Iterable<ChatMessageMention>)
      ) as NullableOption<ChatMessageMention[]>;
    });
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private readonly onMessageDeleted = (message: ChatMessage) => {
    void this._cache.deleteMessage(this.chatId, message);
    this.notifyStateChange((draft: GraphChatClient) => {
      const draftMessage = draft.messages.find(m => m.messageId === message.id) as AcsChatMessage;
      // TODO: confirm if we should show the deleted content message in all cases or only when the message was deleted by the current user
      if (draftMessage) this.setDeletedContent(draftMessage);
    });
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private readonly onMessageEdited = async (message: ChatMessage) => {
    await this._cache.cacheMessage(this._chatId, message);
    const messageConversion = this.convertChatMessage(message);
    this.updateMessages(messageConversion.currentValue);
    if (messageConversion.futureValue) {
      const eventualState = await messageConversion.futureValue;
      this.updateMessages(eventualState);
    }
  };

  private readonly checkForMissedMessages = async () => {
    const messages: MessageCollection = await loadChatThread(this.graph, this.chatId, this._messagesPerCall);
    const messageConversions = messages.value
      // trying to filter out messages on the graph request causes a 400
      // deleted messages are returned as messages with no content, which we can't filter on the graph request
      // so we filter them out here
      // Violating DLP returns content as empty BUT with policyViolation set
      .filter(m => m.body?.content || (!m.body?.content && m?.policyViolation))
      // This gives us both current and eventual values for each message
      .map(m => this.convertChatMessage(m));

    // update the state with the current values
    const currentValueMessages: GraphChatMessage[] = [];
    messageConversions
      .map(m => m.currentValue)
      // need to use a reduce here to filter out undefined values in a way that TypeScript understands
      .reduce((acc, val) => {
        if (val) acc.push(val);
        return acc;
      }, currentValueMessages);
    currentValueMessages.forEach(m => this.updateMessages(m));
    const futureMessages = messageConversions.filter(m => m.futureValue).map(m => m.futureValue);
    // if there are eventual future values, wait for them to resolve and update the state
    if (futureMessages.length > 0) {
      (await Promise.all(futureMessages)).forEach(m => {
        this.updateMessages(m);
      });
    }
    const hasOverlapWithExistingMessages = messages.value.some(m =>
      this._state.messages.find(sm => sm.messageId === m.id)
    );
    if (!hasOverlapWithExistingMessages) {
      // TODO handle the case where there were a lot of missed messages and we ned to get the next page of messages.
      // This is not a common case, but we should handle it.
    }
    log('checked for missed messages');
  };

  private readonly onChatNotificationsSubscribed = (resource: string): void => {
    if (resource.includes(`/${this.chatId}/`) && resource.includes('/messages')) {
      void this.checkForMissedMessages();
    } else {
      // better clean this up as we don't want to be listening to events for other chats
    }
  };

  private readonly onChatPropertiesUpdated = (chat: Chat): void => {
    this._chat = chat;
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.chat = chat;
    });
  };

  private readonly onParticipantAdded = (added: AadUserConversationMember): void => {
    this.notifyStateChange((draft: GraphChatClient) => {
      if (!draft.participants.find(p => p.id === added.id)) {
        draft.participants.push(added);
      }
    });
  };

  private readonly onParticipantRemoved = (added: AadUserConversationMember): void => {
    if (added.id) this.removeParticipantFromState(added.id);
  };

  /**
   * Update the state with given message either replacing an existing message matching on the id or adding to the list
   *
   * @private
   * @param {(GraphChatMessage)} [message]
   * @return {*}
   * @memberof StatefulGraphChatClient
   */
  private updateMessages(message?: GraphChatMessage) {
    if (!message) return;
    this.notifyStateChange((draft: GraphChatClient) => {
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

  private removeParticipantFromState(membershipId: string): void {
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.participants.findIndex(p => p.id === membershipId);
      if (index !== -1) {
        draft.participants.splice(index, 1);
      }
    });
  }

  private readonly addChatMembers = async (userIds: string[], history?: Date): Promise<void> => {
    await addChatMembers(this.graph, this.chatId, userIds, history);
  };

  /**
   * Removes the given member from the chat
   *
   * @param membershpId {string} id of the user for the chat,
   * this is the id of the chat member resource in the graph, i.e. member.id, and NOT the user's id
   * @returns {Promise<void>}
   */
  private readonly removeChatMember = async (membershpId: string): Promise<void> => {
    if (!membershpId) return;
    const isPresent = this._chat?.members?.findIndex(m => m.id === membershpId) ?? -1;
    if (isPresent === -1) return;
    await removeChatMember(this.graph, this.chatId, membershpId);
    this.removeParticipantFromState(membershpId);
  };

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
    content = rewriteEmojiContentToHTML(content);
    // Handle any mentions in the content
    content = this.updateMentionsContent(content);

    const imageMatch = this.graphImageMatch(content ?? '');
    if (imageMatch) {
      // if the message contains an image, we need to fetch the image and replace the placeholder
      result = this.processMessageContent(graphMessage, currentUser);
    } else {
      result.currentValue = this.buildAcsMessage(graphMessage, currentUser, messageId, content);
    }
    return result;
  }

  /**
   * Teams mentions are in the pattern <at id="1">User</at>. This replacement
   * changes the mentions pattern to <msft-mention id="1">User</msft-mention>
   * which will trigger the `mentionOptions` prop to be called in MessageThread.
   *
   * @param content is the message with mentions.
   * @returns string with replaced mention parts.
   */
  private updateMentionsContent(content: string): string {
    const msftMention = `<msft-mention id="$1">$2</msft-mention>`;
    const atRegex = /<at\sid="(\d+)">([a-z0-9_.-\s]+)<\/at>/gim;
    content = content
      .replace(/&nbsp;<at/gim, '<at')
      .replace(/at>&nbsp;/gim, 'at>')
      .replace(atRegex, msftMention);
    return content;
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

  private readonly renameChat = async (topic: string | null): Promise<void> => {
    await updateChatTopic(this.graph, this.chatId, topic);
    this.notifyStateChange(() => void (this._chat = { ...this._chat, ...{ topic } }));
  };

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', (message: ChatMessage) => void this.onMessageReceived(message));
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', (message: ChatMessage) => void this.onMessageEdited(message));
    this._eventEmitter.on('notificationsSubscribedForResource', this.onChatNotificationsSubscribed);
    this._eventEmitter.on('chatThreadPropertiesUpdated', this.onChatPropertiesUpdated);
    // TODO: add define how to handle chat deletion
    // this._eventEmitter.on('chatThreadDeleted', this.onChatDeleted);
    this._eventEmitter.on('participantAdded', this.onParticipantAdded);
    this._eventEmitter.on('participantRemoved', this.onParticipantRemoved);
    this._eventEmitter.on('graphNotificationClientError', this.onGraphNotificationClientError);
  }

  /**
   * Handles the GraphNotificationClient emitted errors.
   */
  private readonly onGraphNotificationClientError = (e: Error) => {
    this.updateErrorState(e);
  };

  /**
   * Emits the error state to the stateful client.
   */
  private readonly updateErrorState = (e: Error) => {
    if (e?.message) {
      error(e?.message);
    } else {
      error('an unknown error occurred', e);
    }
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'error';
      draft.activeErrorMessages = [e];
      draft.chat = { topic: 'Unknown' };
    });
  };

  /**
   * Provided the graph instance for the component with the correct SDK version decoration
   *
   * @readonly
   * @private
   * @type {IGraph}
   * @memberof StatefulGraphChatClient
   */
  private get graph(): IGraph {
    return graph('mgt-chat');
  }

  private readonly _initialState: GraphChatClient = {
    status: 'initial',
    userId: '',
    messages: [],
    mentions: [],
    participants: [],
    get participantCount() {
      return this.participants?.length || 0;
    },
    disableEditing: false,
    numberOfChatMessagesToReload: this._messagesPerCall,
    onDeleteMessage: this.deleteMessage,
    onSendMessage: this.sendMessage,
    onUpdateMessage: this.updateMessage,
    onAddChatMembers: this.addChatMembers,
    onRemoveChatMember: this.removeChatMember,
    onRenameChat: this.renameChat,
    activeErrorMessages: [],
    chat: this._chat
  };
  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  private _state: GraphChatClient = { ...this._initialState };
}

export { StatefulGraphChatClient };
