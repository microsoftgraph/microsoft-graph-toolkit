/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { useEffect, useState } from 'react';
import { v4 as uuid } from 'uuid';
import { StatefulGraphChatListClient } from './StatefulGraphChatListClient';
import { log } from '@microsoft/mgt-element';

/**
 * The key name to use for storing the sessionId in session storage
 */
const keyName = 'mgt-chat-session-id';

/**
 * reads a string from session storage, or if there is no string for the keyName, generate a new uuid and place in storage
 */
const getOrGenerateSessionId = () => {
  const value = sessionStorage.getItem(keyName);

  if (value) {
    return value;
  } else {
    const newValue = uuid();
    sessionStorage.setItem(keyName, newValue);
    return newValue;
  }
};

/**
 * Provides a stable sessionId for the lifetime of the browser tab.
 * @returns a string that is either read from session storage or generated and placed in session storage
 */
const useSessionId = (): string => {
  // when a function is passed to useState, it is only invoked on the first render
  const [sessionId] = useState<string>(getOrGenerateSessionId);

  return sessionId;
};

/**
 * Custom hook to abstract the creation of a stateful graph chat list client.
 * @returns {StatefulGraphChatClient} a stateful graph chat client that is subscribed to the given chatId
 */
export const useGraphChatListClient = (): StatefulGraphChatListClient => {
  const sessionId = useSessionId();
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
