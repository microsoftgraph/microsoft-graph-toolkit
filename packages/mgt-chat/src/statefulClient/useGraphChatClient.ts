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
 * Provides a stable sessionId for the lifetime of the browser tab.
 * @returns a string that is either read from session storage or generated and placed in session storage
 */
const useSessionId = (): string => {
  const [sessionId] = useState<string>(() => uuid());

  return sessionId;
};

/**
 * Custom hook to abstract the creation of a stateful graph chat client.
 * @param {string} chatId the current chatId to be rendered
 * @param {React.EffectCallback} pre an optional useEffect function to be run before other useEffects
 * @param {React.EffectCallback} post an optional useEffect function to be run after other useEffects
 * @returns {StatefulGraphChatClient} a stateful graph chat client that is subscribed to the given chatId
 */
export const useGraphChatClient = (
  chatId: string,
  pre: React.EffectCallback | undefined = undefined,
  post: React.EffectCallback | undefined = undefined
): StatefulGraphChatClient => {
  const sessionId = useSessionId();
  const [chatClient] = useState<StatefulGraphChatClient>(() => new StatefulGraphChatClient());

  // pre
  // NOTE: we need the registration of state handlers before changing state
  useEffect(() => {
    if (pre) pre();
  }, [chatClient, pre]);

  // when chatId or sessionId changes this effect subscribes or unsubscribes
  // the component to/from web socket based notifications for the given chatId
  useEffect(() => {
    // we must have both a chatId & sessionId to subscribe.
    if (chatId && sessionId) chatClient.subscribeToChat(chatId, sessionId);
    else chatClient.setStatus('no chat id');
  }, [chatId, sessionId, chatClient]);

  // Returns a cleanup function to call tearDown on the chatClient
  // This allows us to clean up when the consuming component is being unmounted from the DOM
  useEffect(() => {
    return () => {
      log('invoked clean up effect');
      chatClient.tearDown();
    };
  }, [chatClient]);

  // todo: take out pre/post and move useeffect on line 46 back to Chat.
  // post
  useEffect(() => {
    if (post) post();
  }, [chatClient, post]);

  return chatClient;
};
