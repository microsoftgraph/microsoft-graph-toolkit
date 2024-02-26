/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { useEffect, useState } from 'react';
import { StatefulGraphChatClient } from './StatefulGraphChatClient';
import { log } from '@microsoft/mgt-element';

/**
 * Custom hook to abstract the creation of a stateful graph chat client.
 * @param {string} chatId the current chatId to be rendered
 * @param {React.EffectCallback} pre an optional useEffect function to be run before other useEffects
 * @param {React.EffectCallback} post an optional useEffect function to be run after other useEffects
 * @returns {StatefulGraphChatClient} a stateful graph chat client that is subscribed to the given chatId
 */
export const useGraphChatClient = (): StatefulGraphChatClient => {
  const [chatClient] = useState<StatefulGraphChatClient>(() => new StatefulGraphChatClient());

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
