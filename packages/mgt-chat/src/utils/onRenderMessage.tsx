/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MessageProps, MessageRenderer } from '@azure/communication-react';
import produce from 'immer';
import React from 'react';
import { isChatMessage, isGraphChatMessage } from './types';

import { renderToString } from 'react-dom/server';
import MgtMessageContainer from '../components/ChatContainer/ChatContainer';
import MgtAdaptiveCard from '../components/MgtAdaptiveCard/MgtAdaptiveCard';
import UnsupportedContent from '../components/UnsupportedContent/UnsupportedContent';
/**
 * This is a _dirty_ hack to bundle the teams light CSS used in adaptive cards
 * designer.
 * THOUGHT: import this on demand based on the set theme?
 */
import 'adaptivecards-designer/dist/containers/teams-container-light.css';

/**
 * Renders the preferred content depending on whether it is supported.
 *
 * @param messageProps final message values from the state.
 * @param defaultOnRender default component to render content.
 * @returns
 */
const onRenderMessage = (messageProps: MessageProps, defaultOnRender?: MessageRenderer) => {
  const message = messageProps?.message;
  if (isGraphChatMessage(message)) {
    const attachments = message?.attachments ?? [];

    if (message?.hasUnsupportedContent) {
      const unsupportedContentComponent = <UnsupportedContent targetUrl={message.rawChatUrl} />;
      messageProps = produce(messageProps, (draft: MessageProps) => {
        if (isChatMessage(draft.message)) {
          draft.message.content = renderToString(unsupportedContentComponent);
        }
      });
    } else if (attachments.length) {
      return (
        <MgtAdaptiveCard attachments={attachments} defaultOnRender={defaultOnRender} messageProps={messageProps} />
      );
    }
  }

  return <MgtMessageContainer messageProps={messageProps} defaultOnRender={defaultOnRender} />;
  // return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};

export { onRenderMessage };
