import { MessageProps, MessageRenderer } from '@azure/communication-react';
import produce from 'immer';
import React from 'react';
import { renderToString } from 'react-dom/server';
import UnsupportedContent from '../components/UnsupportedContent/UnsupportedContent';
import { isChatMessage, isGraphChatMessage } from '../utils/types';

/**
 * Renders the preferred content depending on whether it is supported.
 *
 * @param messageProps final message values from the state.
 * @param defaultOnRender default component to render content.
 * @returns
 */
const onRenderMessage = (messageProps: MessageProps, defaultOnRender?: MessageRenderer) => {
  const message = messageProps?.message;
  if (isGraphChatMessage(message) && message?.hasUnsupportedContent) {
    const unsupportedContentComponent = <UnsupportedContent targetUrl={message.rawChatUrl} />;
    messageProps = produce(messageProps, (draft: MessageProps) => {
      if (isChatMessage(draft.message)) {
        draft.message.content = renderToString(unsupportedContentComponent);
      }
    });
  }

  return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};
export { onRenderMessage };
