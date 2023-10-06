/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MessageProps, MessageRenderer } from '@azure/communication-react';
import * as AdaptiveCards from 'adaptivecards';
import MarkdownIt from 'markdown-it';
import React from 'react';
import { isGraphChatMessage, isActionOpenUrl } from './types';
import {
  IAdaptiveCard,
  ISubmitAction,
  IOpenUrlAction,
  IShowCardAction,
  IExecuteAction
} from 'adaptivecards/lib/schema';
import { ChatMessageAttachment } from '@microsoft/microsoft-graph-types';

type IAction = ISubmitAction | IOpenUrlAction | IShowCardAction | IExecuteAction;

/**
 * Renders the preferred content depending on whether it is supported.
 *
 * @param messageProps final message values from the state.
 * @param defaultOnRender default component to render content.
 * @returns
 */
const onRenderMessage = (messageProps: MessageProps, defaultOnRender?: MessageRenderer) => {
  const message = messageProps?.message;
  if (isGraphChatMessage(message) && message?.attachments) {
    const attachments = message?.attachments;
    const render = handleAttachments(attachments, messageProps, defaultOnRender);
    // NOTES:
    // Rendering the card as a JSX element does not render the sender name and the
    // date the message was send. BUT it's the only way to get the buttons to
    // respond to click events from actions.
    //
    // Rendering the card as an HTML string and updating the content gets the
    // UI as expected but we lose the ability to trigger click events on the
    // rendered buttons.
    // messageProps = produce(messageProps, (draft: MessageProps) => {
    //   if (isChatMessage(draft.message)) {
    //     draft.message.content = renderToString(card);
    //   }
    // });
    return render;
  }
  return defaultOnRender ? defaultOnRender(messageProps) : <></>;
};

const handleAttachments = (
  attachments: ChatMessageAttachment[],
  messageProps: MessageProps,
  defaultOnRender?: MessageRenderer
): JSX.Element => {
  const unsupportedCards = [
    'application/vnd.microsoft.card.list',
    'application/vnd.microsoft.card.hero',
    'application/vnd.microsoft.card.o365connector',
    'application/vnd.microsoft.card.receipt',
    'application/vnd.microsoft.card.thumbnail'
  ];

  let render: JSX.Element = defaultOnRender ? defaultOnRender(messageProps) : <></>;

  // Normally it's just a single card per message. Looping through all attachments
  // and only processing those with an adaptive card content type.
  for (const attachment of attachments) {
    const contentType = attachment?.contentType ?? '';
    if (contentType === 'application/vnd.microsoft.card.adaptive') {
      render = handleAdaptiveCards(attachment);
    } else if (unsupportedCards.includes(contentType)) {
      // TODO: update with the correct component
      render = <p>Unsupported content type</p>;
    }
  }
  return render;
};

const handleAdaptiveCards = (attachment: ChatMessageAttachment): JSX.Element => {
  /* eslint-disable
    @typescript-eslint/no-unsafe-assignment,
    @typescript-eslint/no-unsafe-member-access,
    @typescript-eslint/no-unsafe-call, react-hooks/rules-of-hooks */
  const adaptiveCard = new AdaptiveCards.AdaptiveCard();
  const adaptiveCardContentString: string = attachment?.content ?? '';
  const adaptiveCardContent = JSON.parse(adaptiveCardContentString) as IAdaptiveCard;
  const hasSchema = Object.keys(adaptiveCardContent).includes('$schema');

  // Check if the actions property has OpenUrl actions only
  const actions = adaptiveCardContent?.actions?.filter(ac => ac.type === 'Action.OpenUrl');
  if (actions) {
    // Update actions to only Action.OpenUrl actions.
    adaptiveCardContent.actions = actions;
  }

  // Check if the body has actionSet actions and filter for OpenUrl only
  const actionSetArray = adaptiveCardContent?.body?.filter(ac => Object.values(ac).includes('ActionSet'));
  let hasInnerActions = false; // flag to ensure we render with inner actions
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
        hasInnerActions = true;
      }
    }
  }

  if (hasSchema && (actions || hasInnerActions)) {
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
    const renderedCard = adaptiveCard.render() as HTMLElement;
    // BUG: this renders the card 3+ times. Better way to render a card that
    // will respond to click events?
    return <div ref={r => r?.appendChild(renderedCard)}></div>;
  }
  // TODO: update with the correct component
  return <p>Unsupported card content</p>;
};

export { onRenderMessage };
