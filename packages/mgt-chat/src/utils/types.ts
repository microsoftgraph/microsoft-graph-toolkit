/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ChatMessage, Message } from '@azure/communication-react';
import { GraphChatMessage } from 'src/statefulClient/StatefulGraphChatClient';
import { Action, OpenUrlAction } from 'adaptivecards';

/**
 * A typeguard to get the ChatMessage type
 * @param msg of Message
 * @returns {ChatMessage}
 */
export const isChatMessage = (msg: Message): msg is ChatMessage => {
  return 'content' in msg;
};

/**
 * A typeguard to get the GraphChatMessage type
 * @param msg of Message
 * @returns {GraphChatMessage}
 */
export const isGraphChatMessage = (msg: Message): msg is GraphChatMessage => {
  return 'content' in msg && 'hasUnsupportedContent' in msg && 'rawChatUrl' in msg && 'attachments' in msg;
};

/**
 * A typeguard to get the OpenUrlAction type
 * @param o of OpenUrlAction
 * @returns {OpenUrlAction}
 */
export const isActionOpenUrl = (o: Action): o is OpenUrlAction => {
  return 'url' in o;
};
