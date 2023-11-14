/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React from 'react';
import { MessageProps, MessageRenderer } from '@azure/communication-react';
import { getRelativeDisplayDate } from '@microsoft/mgt-components';
import { messageContainer } from '../../utils/messageContainer';
import { isChatMessage } from '../../utils/types';

interface MgtMessageContainerProps {
  messageProps: MessageProps;
  defaultOnRender?: MessageRenderer;
}

const MgtMessageContainer = ({ messageProps, defaultOnRender }: MgtMessageContainerProps) => {
  if (isChatMessage(messageProps.message)) {
    const author = messageProps.message?.senderDisplayName ?? '';
    const timestamp = getRelativeDisplayDate(new Date(messageProps.message.createdOn));
    const details = messageProps.message?.status ?? '';
    const body: string = messageProps.message?.content ?? '';
    const contentType = messageProps.message.contentType;
    const Container = messageContainer(messageProps.message);

    switch (contentType) {
      case 'text':
        return (
          <Container author={author} timestamp={timestamp}>
            {body}
          </Container>
        );
      case 'html':
      case 'richtext/html': {
        const html = { __html: body };
        return (
          <Container author={author} timestamp={timestamp}>
            <div dangerouslySetInnerHTML={html}></div>
          </Container>
        );
      }
      default:
        return defaultOnRender ? defaultOnRender(messageProps) : <></>;
    }
  }
  return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};

export default MgtMessageContainer;
