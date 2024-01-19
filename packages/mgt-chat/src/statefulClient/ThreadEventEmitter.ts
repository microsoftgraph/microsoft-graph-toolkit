/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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
  | 'disconnected'
  | 'connected'
  | 'notificationsSubscribedForResource';

export class ThreadEventEmitter {
  private readonly emitter: EventEmitter = new EventEmitter();

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
  notificationsSubscribedForResource(resouce: string) {
    this.emitter.emit('notificationsSubscribedForResource', resouce);
  }
  disconnected() {
    this.emitter.emit('disconnected');
  }
  connected() {
    this.emitter.emit('connected');
  }
}
