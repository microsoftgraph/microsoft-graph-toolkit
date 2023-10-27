/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { useEffect, useState } from 'react';
import { v4 as uuid } from 'uuid';
import { StatefulGraphChatClient } from './StatefulGraphChatClient';
import { log } from '@microsoft/mgt-element';

export const useGraphChatClient = (chatId: string): StatefulGraphChatClient => {
  const [sessionId, setSessionId] = useState<string | undefined>();
  const [chatClient] = useState<StatefulGraphChatClient>(() => new StatefulGraphChatClient());
  // generate a new sessionId when the chatId changes
  useEffect(() => {
    setSessionId(uuid());
  }, [chatId]);

  // when chatId or sessionId changes this effect subscribes or unsubscribes
  // the component to/from web socket based notifications for the given chatId
  useEffect(() => {
    // we must have both a chatId & sessionId to subscribe.
    if (chatId && sessionId) chatClient.subscribeToChat(chatId, sessionId);
  }, [chatId, sessionId, chatClient]);

  // Returns a cleanup function to call tearDown on the chatClient
  // This allows us to clean up when the consuming component is being unmounted from the DOM
  useEffect(() => {
    return () => {
      log('invoked clean up effect');
      void chatClient.tearDown();
    };
  }, [chatClient]);

  return chatClient;
};
