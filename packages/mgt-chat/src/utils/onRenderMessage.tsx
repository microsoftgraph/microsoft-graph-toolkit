/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Message, MessageProps, MessageRenderer } from '@azure/communication-react';
import { Action, AdaptiveCard, IMarkdownProcessingResult } from 'adaptivecards';
import MarkdownIt from 'markdown-it';
import React, { useEffect, useRef } from 'react';
import { isGraphChatMessage, isActionOpenUrl, isChatMessage } from './types';
import {
  IAdaptiveCard,
  ISubmitAction,
  IOpenUrlAction,
  IShowCardAction,
  IExecuteAction
} from 'adaptivecards/lib/schema';
import { ChatMessageAttachment } from '@microsoft/microsoft-graph-types';
import produce from 'immer';
import { ChatMessage, ChatMyMessage } from '@fluentui-contrib/react-chat';

import { renderToString } from 'react-dom/server';
import UnsupportedContent from '../components/UnsupportedContent/UnsupportedContent';
import { onDisplayDateTimeString } from './displayDates';

type IAction = ISubmitAction | IOpenUrlAction | IShowCardAction | IExecuteAction;

/**
 * Props for an adaptive card message.s
 */
interface MgtAdaptiveCardProps {
  attachments: ChatMessageAttachment[];
  defaultOnRender?: (props: MessageProps) => JSX.Element;
  messageProps: MessageProps;
}

/**
 * Render an adaptive card from the attachments
 */
const MgtAdaptiveCard = (msg: MgtAdaptiveCardProps) => {
  const cardRef = useRef<HTMLDivElement | null>(null);
  const attachments = msg.attachments;
  const adaptiveCardAttachments = getAdaptiveCardAttachments(attachments);
  useEffect(() => {
    if (adaptiveCardAttachments.length) {
      const cardElement = cardRef?.current;
      // Remove all children before appending the attachment elements
      while (cardElement?.firstChild) cardElement.removeChild(cardElement?.lastChild as Node);
      for (const attachment of adaptiveCardAttachments) {
        const cardHtmlElement = getHtmlElementFromAttachment(attachment);
        cardElement?.appendChild(cardHtmlElement!);
      }
    }
  }, [cardRef, adaptiveCardAttachments]);
  const defaultOnRender = msg?.defaultOnRender;
  const messageProps = msg.messageProps;
  const defaultRender = defaultOnRender ? defaultOnRender(messageProps) : <></>;
  const Container = messageContainer(msg.messageProps.message);
  const author = isChatMessage(msg.messageProps.message) ? msg.messageProps.message?.senderDisplayName : '';
  const timestamp = onDisplayDateTimeString(msg.messageProps.message.createdOn);

  return adaptiveCardAttachments.length ? (
    <Container author={author} timestamp={timestamp}>
      <div ref={cardRef}></div>
    </Container>
  ) : (
    defaultRender
  );
};

/**
 * Determine which message container to render. By default use the ChatMessage.
 * @param msg is the Message
 */
const messageContainer = (msg: Message) => {
  if (isChatMessage(msg) && msg?.mine) {
    return ChatMyMessage;
  }
  return ChatMessage;
};

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
  return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};

/**
 * Filters out the adaptive card attachments.
 * @param attachments
 * @returns
 */
const getAdaptiveCardAttachments = (attachments: ChatMessageAttachment[]): ChatMessageAttachment[] => {
  const cardAttachments: ChatMessageAttachment[] = [];
  for (const att of attachments) {
    const contentType = att?.contentType ?? '';
    if (contentType === 'application/vnd.microsoft.card.adaptive') {
      cardAttachments.push(att);
    }
  }
  return cardAttachments;
};

/**
 * Process the attachment object and return an HTMLElement or nothing.
 * @param attachment
 * @returns
 */
const getHtmlElementFromAttachment = (attachment: ChatMessageAttachment | undefined): HTMLElement | undefined => {
  const adaptiveCard = new AdaptiveCard();
  const adaptiveCardContentString: string = attachment?.content ?? '';
  const adaptiveCardContent = JSON.parse(adaptiveCardContentString) as IAdaptiveCard;

  // Check if the actions property has OpenUrl actions only
  const actions = adaptiveCardContent?.actions?.filter(ac => ac.type === 'Action.OpenUrl');
  if (actions) {
    // Update actions to only Action.OpenUrl actions.
    adaptiveCardContent.actions = actions;
  }

  // Check if the body has actionSet actions and filter for OpenUrl only
  const actionSetArray = adaptiveCardContent?.body?.filter(ac => Object.values(ac).includes('ActionSet'));
  if (actionSetArray) {
    const finalInnerActions = [];
    for (const actionSet of actionSetArray) {
      const innerActions = actionSet?.actions as IAction[];
      const valid = innerActions?.filter(ac => ac?.type === 'Action.OpenUrl');
      if (valid) finalInnerActions.push(...valid);
    }

    for (const b of adaptiveCardContent?.body ?? []) {
      if (Object.values(b).includes('ActionSet')) {
        b.actions = finalInnerActions;
      }
    }
  }

  // markdown support
  AdaptiveCard.onProcessMarkdown = (text: string, result: IMarkdownProcessingResult) => {
    const md = new MarkdownIt();
    result.outputHtml = md.render(text);
    result.didProcess = true;
  };

  adaptiveCard.parse(adaptiveCardContent);
  adaptiveCard.onExecuteAction = (action: Action) => {
    if (isActionOpenUrl(action)) {
      const url: string = action?.url ?? '';
      window.open(url, '_blank', 'noopener,noreferrer');
    }
  };
  return adaptiveCard.render();
};

export { onRenderMessage };
