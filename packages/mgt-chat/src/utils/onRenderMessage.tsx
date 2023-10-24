/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MessageProps, MessageRenderer } from '@azure/communication-react';
import * as AdaptiveCards from 'adaptivecards';
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
import { renderToString } from 'react-dom/server';
import UnsupportedContent from '../components/UnsupportedContent/UnsupportedContent';

type IAction = ISubmitAction | IOpenUrlAction | IShowCardAction | IExecuteAction;

/**
 * Props for an adaptive card message.s
 */
interface MGTAdaptiveCardProps {
  attachments: ChatMessageAttachment[];
  defaultOnRender?: (props: MessageProps) => JSX.Element;
  messageProps: MessageProps;
}

/**
 * Render an adaptive card from the attachments
 */
const MGTAdaptiveCard = (msg: MGTAdaptiveCardProps) => {
  const cardRef = useRef<HTMLDivElement | null>(null);
  const attachments = msg.attachments;
  const adaptiveCardAttachment = getAdaptiveCardAttachment(attachments);
  useEffect(() => {
    if (adaptiveCardAttachment) {
      const cardHtmlElement = getHtmlElementFromAttachment(adaptiveCardAttachment);
      cardRef?.current?.appendChild(cardHtmlElement!);
    }
  }, [cardRef, adaptiveCardAttachment]);
  const defaultOnRender = msg?.defaultOnRender;
  const messageProps = msg.messageProps;
  const defaultRender = defaultOnRender ? defaultOnRender(messageProps) : <></>;
  return adaptiveCardAttachment ? <div ref={cardRef}></div> : defaultRender;
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
        <MGTAdaptiveCard attachments={attachments} defaultOnRender={defaultOnRender} messageProps={messageProps} />
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
const getAdaptiveCardAttachment = (attachments: ChatMessageAttachment[]): ChatMessageAttachment | undefined => {
  for (const att of attachments) {
    const contentType = att?.contentType ?? '';
    if (contentType === 'application/vnd.microsoft.card.adaptive') {
      return att;
    }
  }
  return;
};

/**
 * Process the attachment object and return an HTMLElement or nothing.
 * @param attachment
 * @returns
 */
const getHtmlElementFromAttachment = (attachment: ChatMessageAttachment | undefined): HTMLElement | undefined => {
  const adaptiveCard = new AdaptiveCards.AdaptiveCard();
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
  AdaptiveCards.AdaptiveCard.onProcessMarkdown = (text: string, result: AdaptiveCards.IMarkdownProcessingResult) => {
    const md = new MarkdownIt();
    result.outputHtml = md.render(text);
    result.didProcess = true;
  };

  adaptiveCard.parse(adaptiveCardContent);
  adaptiveCard.onExecuteAction = (action: AdaptiveCards.Action) => {
    if (isActionOpenUrl(action)) {
      const url: string = action?.url ?? '';
      window.open(url, '_blank', 'noopener,noreferrer');
    }
  };
  return adaptiveCard.render();
};

export { onRenderMessage };
