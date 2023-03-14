import {
  MessageThreadProps,
  SendBoxProps,
  ChatMessage as ACSChatMessage,
  ErrorBarProps
} from '@azure/communication-react';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';
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
  updateChatMessage
} from './graph.chat';
import { graphChatMessageToACSChatMessage } from './acs.chat';

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
  Pick<ErrorBarProps, 'activeErrorMessages'>;

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

const MessageCreatedComparator = (a: ACSChatMessage, b: ACSChatMessage) =>
  a.createdOn.getTime() - b.createdOn.getTime();

class StatefulGraphChatClient implements StatefulClient<GraphChatClient> {
  private _subscribers: ((state: GraphChatClient) => void)[] = [];
  private _messagesPerCall = 5;
  private _nextLink: string = '';
  private _chat?: Chat = undefined;
  private _userDisplayName: string = '';

  constructor() {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
  }

  public onStateChange(handler: (state: GraphChatClient) => void): void {
    if (!this._subscribers.includes(handler)) {
      this._subscribers.push(handler);
    }
  }

  public offStateChange(handler: (state: GraphChatClient) => void): void {
    const index = this._subscribers.indexOf(handler);
    if (index !== -1) {
      this._subscribers = this._subscribers.splice(index, 1);
    }
  }

  private notifyStateChange(recipe: (draft: GraphChatClient) => void) {
    this._state = produce(this._state, recipe);
    this._subscribers.forEach(handler => handler(this._state));
  }

  public getState(): GraphChatClient {
    return this._state;
  }

  private onStateChanged = (e: LoginChangedEvent) => {
    switch (Providers.globalProvider.state) {
      case ProviderState.SignedIn:
        // update userId and displayName
        this.updateUserInfo();
        // load messages?
        // configure subscriptions
        // emit new state;
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

  private async updateFollowedChat() {
    // Subscribe to notifications for messages
    // Subscribe to notifications for the chat itself?
    // load Chat Data
    this._chat = await loadChat(this.graph, this._chatId);
    const messages: MessageCollection = await loadChatThread(this.graph, this._chatId, this._messagesPerCall);
    this._nextLink = messages.nextLink;
    // Allow messages to be loaded via the loadMoreMessages callback
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.participantCount = this._chat?.members?.length || 0;
      draft.messages = messages.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.messageType === 'message' && m.body?.content)
        .map(m => graphChatMessageToACSChatMessage(m, this._userId));
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
    });
  }

  private loadMoreMessages = async () => {
    if (!this._nextLink) {
      return true;
    }
    const result: MessageCollection = await loadMoreChatMessages(this.graph, this._nextLink);

    this._nextLink = result.nextLink;
    this.notifyStateChange((draft: GraphChatClient) => {
      const nextMessages = result.value
        // trying to filter out system messages on the graph request causes a 400
        // delted messages are returned as messages with no content, which we can't filter on the graph request
        // so we filter them out here
        .filter(m => m.messageType === 'message' && m.body?.content)
        .map(m => graphChatMessageToACSChatMessage(m, this._userId));
      draft.messages = nextMessages.concat(draft.messages as ACSChatMessage[]).sort(MessageCreatedComparator);
      draft.onLoadPreviousChatMessages = this._nextLink ? this.loadMoreMessages : undefined;
    });
    // return true when there are no more messages to load
    return !Boolean(this._nextLink);
  };

  private sendMessage = async (content: string) => {
    if (!content) return;

    const pendingId = uuid();

    // add a pending message to the state.
    this.notifyStateChange((draft: GraphChatClient) => {
      const pendingMessage: ACSChatMessage = {
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
        draft.messages.splice(draftIndex, 1, graphChatMessageToACSChatMessage(chat, this._userId));
      });
    } catch (e) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const draftMessage = draft.messages.find(m => m.messageId === pendingId);
        (draftMessage as ACSChatMessage).status = 'failed';
      });
      throw new Error('Failed to send message');
    }
  };

  public deleteMessage = async (messageId: string) => {
    if (!messageId) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as ACSChatMessage;
    // only messages not persisted to graph should have a clientMessageId
    const uncommitted = this._state.messages.find(
      m => (m as ACSChatMessage).clientMessageId === messageId
    ) as ACSChatMessage;
    if (message?.mine) {
      try {
        // uncommitted messages are not persisted to the graph, so don't call graph when deleting them
        if (!uncommitted) {
          await deleteChatMessage(this.graph, this._chatId, messageId);
        }
        this.notifyStateChange((draft: GraphChatClient) => {
          const draftMessage = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          if (draftMessage.clientMessageId) {
            // just remove messages that were not saved to the graph
            draft.messages.splice(draft.messages.indexOf(draftMessage), 1);
          } else {
            // show deleted messages which have been persisted to the graph as deleted in the UI
            draftMessage.content = '<em>This message has been deleted.</em>';
            draftMessage.contentType = 'html';
          }
        });
      } catch (e) {
        //TODO: How do we handle failed deletes?
      }
    }
  };

  public updateMessage = async (messageId: string, content: string) => {
    if (!messageId || !content) return;
    const message = this._state.messages.find(m => m.messageId === messageId) as ACSChatMessage;
    if (message?.mine && message.content) {
      this.notifyStateChange((draft: GraphChatClient) => {
        const updating = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
        if (updating) {
          updating.content = content;
          updating.status = 'sending';
        }
      });
      try {
        await updateChatMessage(this.graph, this._chatId, messageId, content);
        this.notifyStateChange((draft: GraphChatClient) => {
          const updated = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          updated.status = 'delivered';
        });
      } catch (e) {
        this.notifyStateChange((draft: GraphChatClient) => {
          const updating = draft.messages.find(m => m.messageId === messageId) as ACSChatMessage;
          updating.status = 'failed';
        });
        throw new Error('Failed to update message');
      }
    }
  };

  private get graph(): IGraph {
    return Providers.globalProvider.graph.forComponent('mgt-chat');
  }

  private _state: GraphChatClient = {
    userId: '',
    messages: [],
    participantCount: 0,
    disableEditing: false,
    numberOfChatMessagesToReload: this._messagesPerCall,
    onDeleteMessage: this.deleteMessage,
    onSendMessage: this.sendMessage,
    onUpdateMessage: this.updateMessage,
    activeErrorMessages: []
  };
}

export { StatefulGraphChatClient };
