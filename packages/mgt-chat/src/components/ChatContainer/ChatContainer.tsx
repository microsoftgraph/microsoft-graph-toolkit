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
  // TODO: find out how to render emojis
  // <p>This is <emoji id=\"cool\" alt=\"😎\" title=\"Cool\"></emoji></p>
  if (isChatMessage(messageProps.message)) {
    const author = messageProps.message?.senderDisplayName ?? '';
    const timestamp = getRelativeDisplayDate(new Date(messageProps.message.createdOn));
    const details = messageProps.message?.status ?? '';
    const body: string = messageProps.message?.content ?? '';
    const contentType = messageProps.message.contentType;
    const Container = messageContainer(messageProps.message);

    console.log(messageProps);
    switch (contentType) {
      case 'text':
        return (
          <Container author={author} timestamp={timestamp}>
            {body}
          </Container>
        );
      case 'html':
      case 'richtext/html': {
        const bodyContent = processEmojiContent(body);
        const html = { __html: bodyContent };
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

/**
 * Regex to detect and extract emoji alt text
 *
 * Pattern breakdown:
 * (<emoji[^>]+): Captures the opening emoji tag, including any attributes.
 * alt=["'](\w*[^"']*)["']: Matches and captures the "alt" attribute value within single or double quotes. The value can contain word characters but not quotes.
 * (.*[^>]): Captures any remaining text within the opening emoji tag, excluding the closing tag.
 * </emoji>: Matches the closing emoji tag.
 */
const emojiRegex = /(<emoji[^>]+)alt=["'](\w*[^"']*)["'](.*[^>])<\/emoji>/;

const emojiMatch = (messageContent: string): RegExpMatchArray | null => {
  return messageContent.match(emojiRegex);
};

const processEmojiContent = (messageContent: string): string => {
  let result = messageContent;
  let match = emojiMatch(result);
  while (match) {
    result = result.replace(emojiRegex, '$2');
    match = emojiMatch(result);
  }
  return result;
};

export default MgtMessageContainer;
