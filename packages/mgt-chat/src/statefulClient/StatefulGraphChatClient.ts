import { MessageThreadProps, SendBoxProps } from '@azure/communication-react';
import { Chat, ChatMessage } from '@microsoft/microsoft-graph-types';
import { ActiveAccountChanged, IGraph, LoginChangedEvent, Providers, ProviderState } from '@microsoft/mgt-element';
import { produce } from 'immer';
import { loadChat, loadChatThread, loadMoreChatMessages, MessageCollection, sendChatMessage } from './graph.chat';
import { graphChatMessageToACSChatMessage } from './acs.chat';
import { ChatMessage as ACSChatMessage } from '@azure/communication-react';

type GraphChatClient = Pick<
  MessageThreadProps,
  | 'userId'
  | 'messages'
  | 'participantCount'
  | 'disableEditing'
  | 'onLoadPreviousChatMessages'
  | 'numberOfChatMessagesToReload'
> &
  Pick<SendBoxProps, 'onSendMessage'>;

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

  constructor() {
    this.updateCurrentUserId();
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
        this.updateCurrentUserId();
        // update userId
        // load messages
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
    this.updateCurrentUserId();
  };

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
    // send message
    const chat: ChatMessage = await sendChatMessage(this.graph, this._chatId, content);
    // emit new state
    this.notifyStateChange((draft: GraphChatClient) => {
      draft.messages.push(graphChatMessageToACSChatMessage(chat, this._userId));
    });
  };

  private get graph(): IGraph {
    return Providers.globalProvider.graph.forComponent('mgt-chat');
  }

  private _state: GraphChatClient = {
    userId: '',
    messages: [],
    participantCount: 0,
    disableEditing: true,
    numberOfChatMessagesToReload: this._messagesPerCall,
    onSendMessage: this.sendMessage
  };
}

export { StatefulGraphChatClient };
