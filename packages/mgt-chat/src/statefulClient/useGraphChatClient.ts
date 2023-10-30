/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { useState } from 'react';
import { StatefulGraphChatClient } from './StatefulGraphChatClient';

export const useGraphChatClient = (chatId: string): StatefulGraphChatClient => {
  const [chatClient] = useState<StatefulGraphChatClient>(() => new StatefulGraphChatClient());
  chatClient.chatId = chatId;

  return chatClient;
};
