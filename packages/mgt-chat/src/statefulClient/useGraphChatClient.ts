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

/**
 * Custom hook to abstract the creation of a stateful graph chat client.
 * @param {string} chatId the current chatId to be rendered
 * @returns {StatefulGraphChatClient} a stateful graph chat client that is subscribed to the given chatId
 */
export const useGraphChatClient = (chatId: string): StatefulGraphChatClient => {
  const [chatClient] = useState<StatefulGraphChatClient>(() => new StatefulGraphChatClient());

  // when chatId or sessionId changes this effect subscribes or unsubscribes
  // the component to/from web socket based notifications for the given chatId
  useEffect(() => {
    // we must have both a chatId & sessionId to subscribe.
    if (chatId) chatClient.subscribeToChat(chatId);
    else chatClient.setStatus('no chat id');
  }, [chatId, chatClient]);

  // Returns a cleanup function to call tearDown on the chatClient
  // This allows us to clean up when the consuming component is being unmounted from the DOM
  useEffect(() => {
    return () => {
      log('invoked clean up effect');
      chatClient.tearDown();
    };
  }, [chatClient]);

  return chatClient;
};
