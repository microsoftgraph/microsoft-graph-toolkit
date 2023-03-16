// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ChatMessage } from '@azure/communication-react';
import {
  ChatThreadCreatedEvent,
  ChatThreadDeletedEvent,
  ChatThreadPropertiesUpdatedEvent,
  ParticipantsAddedEvent,
  ParticipantsRemovedEvent,
  ReadReceiptReceivedEvent,
  TypingIndicatorReceivedEvent
} from '@azure/communication-signaling';
import { EventEmitter } from 'events';

export type ChatEvent =
  | 'chatMessageReceived'
  | 'chatMessageEdited'
  | 'chatMessageDeleted'
  | 'typingIndicatorReceived'
  | 'readReceiptReceived'
  | 'chatThreadCreated'
  | 'chatThreadDeleted'
  | 'chatThreadPropertiesUpdated'
  | 'participantsAdded'
  | 'participantsRemoved'
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
  chatThreadCreated(e: ChatThreadCreatedEvent) {
    this.emitter.emit('chatThreadCreated', e);
  }
  chatThreadDeleted(e: ChatThreadDeletedEvent) {
    this.emitter.emit('chatThreadDeleted', e);
  }
  chatThreadPropertiesUpdated(e: ChatThreadPropertiesUpdatedEvent) {
    this.emitter.emit('chatThreadPropertiesUpdated', e);
  }
  participantsAdded(e: ParticipantsAddedEvent) {
    this.emitter.emit('participantsAdded', e);
  }
  participantsRemoved(e: ParticipantsRemovedEvent) {
    this.emitter.emit('participantsRemoved', e);
  }

  chatMessageNotificationsSubscribed(messagesResource: string) {
    this.emitter.emit('chatMessageNotificationsSubscribed', messagesResource);
  }
}
