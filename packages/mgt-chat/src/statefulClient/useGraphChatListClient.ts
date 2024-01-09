/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { useEffect, useState } from 'react';
import { StatefulGraphChatListClient } from './StatefulGraphChatListClient';
import { log } from '@microsoft/mgt-element';

/**
 * Custom hook to abstract the creation of a stateful graph chat list client.
 * @returns {StatefulGraphChatClient} a stateful graph chat client that is subscribed to the given chatId
 */
export const useGraphChatListClient = (): StatefulGraphChatListClient => {
  // sessionId is going to be used to lookup subscription and we need to ensure the same subscription is used across tabs.
  const sessionId = 'default';
  const [chatClient] = useState<StatefulGraphChatListClient>(() => new StatefulGraphChatListClient());

  // when chatId or sessionId changes this effect subscribes or unsubscribes
  // the component to/from web socket based notifications for the given chatId
  useEffect(() => {
    // we must have sessionId to subscribe.
    if (sessionId) chatClient.subscribeToUser(sessionId);
  }, [sessionId, chatClient]);

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
