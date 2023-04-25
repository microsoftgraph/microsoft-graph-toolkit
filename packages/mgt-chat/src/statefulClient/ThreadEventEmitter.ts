// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ReadReceiptReceivedEvent, TypingIndicatorReceivedEvent } from '@azure/communication-signaling';
import { AadUserConversationMember, Chat, ChatMessage } from '@microsoft/microsoft-graph-types';
import { EventEmitter } from 'events';

export type ChatEvent =
  | 'chatMessageReceived'
  | 'chatMessageEdited'
  | 'chatMessageDeleted'
  | 'typingIndicatorReceived'
  | 'readReceiptReceived'
  | 'chatThreadDeleted'
  | 'chatThreadPropertiesUpdated'
  | 'participantAdded'
  | 'participantRemoved'
  | 'chatMessageNotificationsSubscribed';

export class ThreadEventEmitter {
  private emitter: EventEmitter = new EventEmitter();

  on(event: ChatEvent, listener: (...args: any[]) => void) {
    this.emitter.on(event, listener);
  }

  off(event: ChatEvent, listener: (...args: any[]) => void) {
    this.emitter.off(event, listener);
  }

  chatMessageReceived(e: ChatMessage) {
    this.emitter.emit('chatMessageReceived', e);
  }
  chatMessageEdited(e: ChatMessage) {
    this.emitter.emit('chatMessageEdited', e);
  }
  chatMessageDeleted(e: ChatMessage) {
    this.emitter.emit('chatMessageDeleted', e);
  }
  typingIndicatorReceived(e: TypingIndicatorReceivedEvent) {
    this.emitter.emit('typingIndicatorReceived', e);
  }
  readReceiptReceived(e: ReadReceiptReceivedEvent) {
    this.emitter.emit('readReceiptReceived', e);
  }
  chatThreadDeleted(e: Chat) {
    this.emitter.emit('chatThreadDeleted', e);
  }
  chatThreadPropertiesUpdated(e: Chat) {
    this.emitter.emit('chatThreadPropertiesUpdated', e);
  }
  participantAdded(e: AadUserConversationMember) {
    this.emitter.emit('participantAdded', e);
  }
  participantRemoved(e: AadUserConversationMember) {
    this.emitter.emit('participantRemoved', e);
  }

  chatMessageNotificationsSubscribed(messagesResource: string) {
    this.emitter.emit('chatMessageNotificationsSubscribed', messagesResource);
  }
}
