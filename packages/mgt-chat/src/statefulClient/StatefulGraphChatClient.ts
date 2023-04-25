import {
  MessageThreadProps,
  SendBoxProps,
  ChatMessage as AcsChatMessage,
  ErrorBarProps,
  SystemMessage,
  ContentSystemMessage
} from '@azure/communication-react';
import {
  AadUserConversationMember,
  Chat,
  ChatMessage,
  MembersDeletedEventMessageDetail
} from '@microsoft/microsoft-graph-types';
import { ActiveAccountChanged, IGraph, LoginChangedEvent, Providers, ProviderState } from '@microsoft/mgt-element';
import { produce } from 'immer';
import { v4 as uuid } from 'uuid';
import {
  deleteChatMessage,
  loadChat,
  loadChatThread,
  loadMoreChatMessages,
  MessageCollection,
  sendChatMessage,
  updateChatMessage,
  removeChatMember,
  addChatMembers
} from './graph.chat';
import { getUserWithPhoto } from '@microsoft/mgt-components';
import { graphChatMessageToAcsChatMessage } from './acs.chat';
import { GraphNotificationClient } from './GraphNotificationClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { IDynamicPerson } from '@microsoft/mgt-react';

type ODataType = {
  '@odata.type': MessageEventType;
};

type GraphChatClient = Pick<
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
  Pick<SendBoxProps, 'onSendMessage'> &
  Pick<ErrorBarProps, 'activeErrorMessages'> & {
    status:
      | 'initial'
      | 'creating server connections'
      | 'subscribing to notifications'
      | 'loading messages'
      | 'ready'
      | 'error';
    chat?: Chat;
  } & {
    participants: AadUserConversationMember[];
    onAddChatMembers: (userIds: string[], history?: Date) => Promise<void>;
    onRemoveChatMember: (membershipId: string) => Promise<void>;
  };

type StatefulClient<T> = {
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
};

type CreatedOn = {
  createdOn: Date;
};

/**
 * Simple object comparator function for sorting by createdOn date
 *
 * @param {CreatedOn} a
 * @param {CreatedOn} b
 */
const MessageCreatedComparator = (a: CreatedOn, b: CreatedOn) => a.createdOn.getTime() - b.createdOn.getTime();

type MessageEventType =
  | '#microsoft.graph.membersAddedEventMessageDetail'
  | '#microsoft.graph.membersDeletedEventMessageDetail';

class StatefulGraphChatClient implements StatefulClient<GraphChatClient> {
  private _notificationClient: GraphNotificationClient;
  private _eventEmitter: ThreadEventEmitter;
  private _subscribers: ((state: GraphChatClient) => void)[] = [];
  private _messagesPerCall = 5;
  private _nextLink: string = '';
  private _chat?: Chat = undefined;
  private _userDisplayName: string = '';

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._notificationClient = new GraphNotificationClient();
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
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
  private onLoginStateChanged = (e: LoginChangedEvent) => {
    switch (Providers.globalProvider.state) {
      case ProviderState.SignedIn:
        // update userId and displayName
        this.updateUserInfo();
        // load messages?
        // configure subscriptions
        // emit new state;
        if (this._chatId) {
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

  private onActiveAccountChanged = (e: ActiveAccountChanged) => {
    this.updateUserInfo();
  };

  private updateUserInfo() {
    this.updateCurrentUserId();
    this.updateCurrentUserName();
  }

  private updateCurrentUserName() {
    this._userDisplayName = Providers.globalProvider.getActiveAccount?.().name || '';
  }

  private updateCurrentUserId() {
    this.userId = Providers.globalProvider.getActiveAccount?.().id.split('.')[0] || '';
  }

  private _userId: string = '';
  private set userId(userId: string) {
    if (this._userId === userId) {
      return;
    }
    this._userId = userId;
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.userId = userId;
    });
  }

  private _chatId: string = '';

  public set chatId(value: string) {
    // take no action if the chatId is the same
    if (this._chatId === value) {
      return;
    }
    this._chatId = value;
    void this.updateFollowedChat();
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatClient
   */
  private async updateFollowedChat() {
    // reset state to initial
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'initial';
      draft.messages = [];
      draft.chat = undefined;
      draft.participants = [];
    });
    // Subscribe to notifications for messages
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'creating server connections';
    });
    // subscribing to notifications will trigger the chatMessageNotificationsSubscribed event
    // this client will then load the chat and messages when that event listener is called
    await this._notificationClient.subscribeToChatNotifications(this._userId, this._chatId, this._eventEmitter, () =>
      this.notifyStateChange((draft: GraphChatClient) => {
        draft.status = 'subscribing to notifications';
      })
    );
  }

  private async loadChatData() {
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.status = 'loading messages';
    });
    this._chat = await loadChat(this.graph, this._chatId);
    const messages: MessageCollection = await loadChatThread(this.graph, this._chatId, this._messagesPerCall);
    this._nextLink = messages.nextLink;

    const initialMessages = await Promise.all(
      messages.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.body?.content)
        .map(m => this.convertChatMessage(m))
    );
    // Allow messages to be loaded via the loadMoreMessages callback
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.participants = this._chat?.members || [];
      draft.participantCount = draft.participants.length;
      draft.messages = initialMessages;
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
      draft.status = this._nextLink ? 'loading messages' : 'ready';
      draft.chat = this._chat;
    });
    // if for some reason the first page only has un-renderable messages, load more
    if (this._nextLink && initialMessages.length < 1) {
      // load more messages
      await this.loadMoreMessages();
    }
  }

  private convertChatMessage(message: ChatMessage): Promise<AcsChatMessage | SystemMessage> {
    switch (message.messageType) {
      case 'message':
        return Promise.resolve(graphChatMessageToAcsChatMessage(message, this._userId));
      case 'unknownFutureValue':
        return this.buildSystemContentMessage(message);
      default:
        throw new Error(`Unknown message type ${message.messageType}`);
    }
  }

  private resolveIcon = (eventType: MessageEventType): string => {
    switch (eventType) {
      case '#microsoft.graph.membersAddedEventMessageDetail':
        return 'add-friend';
      case '#microsoft.graph.membersDeletedEventMessageDetail':
        return 'left-chat';
      default:
        return 'Unknown';
    }
  };

  private async buildSystemContentMessage(message: ChatMessage): Promise<SystemMessage> {
    const eventDetail = message.eventDetail as MembersDeletedEventMessageDetail & ODataType;
    const awaits: Promise<IDynamicPerson>[] = [];
    const initiatorId = eventDetail.initiator?.user?.id;
    // we're using getUserWithPhoto here because we want to tap into the caching that mgt-person uses to cut down on graph calls
    initiatorId && awaits.push(getUserWithPhoto(this.graph, initiatorId));
    for (const m of eventDetail.members ?? []) {
      awaits.push(getUserWithPhoto(this.graph, m.id!));
    }
    const people = await Promise.all(awaits);
    let messageContent = '';
    switch (eventDetail['@odata.type']) {
      case '#microsoft.graph.membersAddedEventMessageDetail':
        messageContent = `${this.getUserName(initiatorId!, people)} added ${eventDetail.members
          ?.map(m => this.getUserName(m.id!, people))
          .join(', ')}`;
        break;
      case '#microsoft.graph.membersDeletedEventMessageDetail':
        messageContent = `${this.getUserName(initiatorId!, people)} removed ${eventDetail.members
          ?.map(m => this.getUserName(m.id!, people))
          .join(', ')}`;
        break;
      default:
        messageContent = `Unknown system message type ${eventDetail['@odata.type']}

detail: ${JSON.stringify(eventDetail)}`;
    }

    const result: ContentSystemMessage = {
      createdOn: message.createdDateTime ? new Date(message.createdDateTime) : new Date(),
      messageId: message.id!,
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
  private loadMoreMessages = async () => {
    if (!this._nextLink) {
      return true;
    }
    const result: MessageCollection = await loadMoreChatMessages(this.graph, this._nextLink);

    const messages = await Promise.all(
      result.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.body?.content)
        .map(m => this.convertChatMessage(m))
    );
    this._nextLink = result.nextLink;
    this.notifyStateChange((draft: GraphChatClient) => {
      const nextMessages = messages;
      draft.messages = nextMessages.concat(draft.messages as AcsChatMessage[]).sort(MessageCreatedComparator);
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
    });
    // return true when there are no more messages to load
    return !Boolean(this._nextLink);
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
      const pendingMessage: AcsChatMessage = {
        clientMessageId: pendingId,
        messageId: pendingId,
        contentType: 'text',
        messageType: 'chat',
        content: content,
        senderDisplayName: this._userDisplayName,
        createdOn: new Date(),
        senderId: this._userId,
        mine: true,
        status: 'sending'
      };
      draft.messages.push(pendingMessage);
    });
    try {
      // send message
      const chat: ChatMessage = await sendChatMessage(this.graph, this._chatId, content);
      // emit new state
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftIndex = draft.messages.findIndex(m => m.messageId === pendingId);
        draft.messages.splice(draftIndex, 1, graphChatMessageToAcsChatMessage(chat, this._userId));
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
  private setDeletedContent = (message: AcsChatMessage) => {
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
          await deleteChatMessage(this.graph, this._chatId, messageId);
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
        //TODO: How do we handle failed deletes?
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
        await updateChatMessage(this.graph, this._chatId, messageId, content);
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
  private onMessageReceived = async (message: ChatMessage) => {
    const acsMessage = await this.convertChatMessage(message);
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.messages.findIndex(m => m.messageId === acsMessage.messageId);
      // this message is not already in thread so just add it
      if (index === -1) {
        // sort to ensure that messages are in the correct order should we get messages out of order
        draft.messages = draft.messages.concat(acsMessage).sort(MessageCreatedComparator);
      } else {
        // replace the existing version of the message with the new one
        draft.messages.splice(index, 1, acsMessage);
      }
    });
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private onMessageDeleted = (message: ChatMessage) => {
    this.notifyStateChange((draft: GraphChatClient) => {
      const draftMessage = draft.messages.find(m => m.messageId === message.id) as AcsChatMessage;
      if (draftMessage) this.setDeletedContent(draftMessage);
    });
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private onMessageEdited = async (message: ChatMessage) => {
    const acsMessage = await this.convertChatMessage(message);
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.messages.findIndex(m => m.messageId === acsMessage.messageId);
      draft.messages.splice(index, 1, acsMessage);
    });
  };

  private onChatNotificationsSubscribed = (messagesResource: string): void => {
    if (messagesResource.includes(`/${this._chatId}/`)) {
      void this.loadChatData();
    } else {
      // better clean this up as we don't want to be listening to events for other chats
    }
  };

  private onChatPropertiesUpdated = async (chat: Chat): Promise<void> => {
    this._chat = chat;
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.chat = chat;
    });
  };

  private onParticipantAdded = (added: AadUserConversationMember): void => {
    this.notifyStateChange((draft: GraphChatClient) => {
      if (!draft.participants.find(p => p.id === added.id)) {
        draft.participants.push(added);
      }
    });
  };

  private onParticipantRemoved = (added: AadUserConversationMember): void => {
    added.id && this.removeParticipantFromState(added.id);
  };

  private removeParticipantFromState(membershipId: string): void {
    this.notifyStateChange((draft: GraphChatClient) => {
      const index = draft.participants.findIndex(p => p.id === membershipId);
      if (index !== -1) {
        draft.participants.splice(index, 1);
      }
    });
  }

  private addChatMembers = async (userIds: string[], history?: Date): Promise<void> => {
    await addChatMembers(this.graph, this._chatId, userIds, history);
  };

  /**
   * Removes the given member from the chat
   *
   * @param membershpId {string} id of the user for the chat,
   * this is the id of the chat member resource in the graph, i.e. member.id, and NOT the user's id
   * @returns {Promise<void>}
   */
  private removeChatMember = async (membershpId: string): Promise<void> => {
    if (!membershpId) return;
    const isPresent = this._chat?.members?.findIndex(m => m.id === membershpId) ?? -1;
    if (isPresent === -1) return;
    await removeChatMember(this.graph, this._chatId, membershpId);
    this.removeParticipantFromState(membershpId);
  };

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', this.onMessageReceived);
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', this.onMessageEdited);
    this._eventEmitter.on('chatMessageNotificationsSubscribed', this.onChatNotificationsSubscribed);
    this._eventEmitter.on('chatThreadPropertiesUpdated', this.onChatPropertiesUpdated);
    // TODO: add define how to handle chat deletion
    // this._eventEmitter.on('chatThreadDeleted', this.onChatDeleted);
    this._eventEmitter.on('participantAdded', this.onParticipantAdded);
    this._eventEmitter.on('participantRemoved', this.onParticipantRemoved);
  }

  /**
   * Provided the graph instance for the component with the correct SDK version decoration
   *
   * @readonly
   * @private
   * @type {IGraph}
   * @memberof StatefulGraphChatClient
   */
  private get graph(): IGraph {
    return Providers.globalProvider.graph.forComponent('mgt-chat');
  }

  private _initialState: GraphChatClient = {
    status: 'initial',
    userId: '',
    messages: [],
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
