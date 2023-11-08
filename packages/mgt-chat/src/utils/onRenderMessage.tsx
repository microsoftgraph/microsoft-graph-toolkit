/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MessageProps, MessageRenderer } from '@azure/communication-react';
import React, { useEffect, useRef } from 'react';
import { isGraphChatMessage, isChatMessage } from './types';
import produce from 'immer';

import { renderToString } from 'react-dom/server';
import UnsupportedContent from '../components/UnsupportedContent/UnsupportedContent';
import MgtAdaptiveCard from '../components/MgtAdaptiveCard/MgtAdaptiveCard';
/**
 * This is a _dirty_ hack to bundle the teams light CSS used in adaptive cards
 * designer.
 * THOUGHT: import this on demand based on the set theme?
 */
import 'adaptivecards-designer/dist/containers/teams-container-light.css';
import { messageContainer } from './messageContainer';
import { onDisplayDateTimeString } from './displayDates';

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

  // return <MgtMessageContainer messageProps={messageProps} defaultOnRender={defaultOnRender} />;
  return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};

interface MgtMessageContainerProps {
  messageProps: MessageProps;
  defaultOnRender: MessageRenderer;
}

const MgtMessageContainer = ({ messageProps, defaultOnRender }: MgtMessageContainerProps) => {
  const bodyRef = useRef<HTMLDivElement>(null);
  const author = isChatMessage(messageProps.message) ? messageProps.message?.senderDisplayName : '';
  const timestamp = onDisplayDateTimeString(messageProps.message.createdOn);
  const details = isChatMessage(messageProps.message) ? messageProps.message?.status : '';
  const body = messageProps.message?.content;

  useEffect(() => {
    if (bodyRef.current) bodyRef.current.innerHTML = body;
  }, [bodyRef, body]);
  const Container = messageContainer(messageProps.message);
  return body ? <Container author={author} timestamp={timestamp} ref={bodyRef} /> : defaultOnRender(messageProps);
};
export { onRenderMessage };
