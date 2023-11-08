/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Message } from '@azure/communication-react';
import { ChatMessage, ChatMyMessage } from '@fluentui-contrib/react-chat';
import { isChatMessage } from './types';
/**
 * Determine which message container to render. By default use the ChatMessage.
 * @param msg is the Message
 */
export const messageContainer = (msg: Message) => {
  if (isChatMessage(msg) && msg?.mine) {
    return ChatMyMessage;
  }
  return ChatMessage;
};
