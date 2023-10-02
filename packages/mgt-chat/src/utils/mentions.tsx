import React from 'react';
import { PersonCardInteraction } from '@microsoft/mgt-components';
import { MgtTemplateProps, Person } from '@microsoft/mgt-react';
import { ChatMessageMention, User } from '@microsoft/microsoft-graph-types';
import { GraphChatClient } from 'src/statefulClient/StatefulGraphChatClient';
import { Mention } from '@azure/communication-react';

export const renderMGTMention = (chatState: GraphChatClient) => {
  return (mention: Mention, defaultRenderer: (mention: Mention) => JSX.Element): JSX.Element => {
    let render: JSX.Element = defaultRenderer(mention);

    const mentions = chatState?.mentions ?? [];
    // TODO: Array.flat() is a new api, update?
    // eslint-disable-next-line @typescript-eslint/no-unnecessary-type-assertion
    const flatMentions = mentions?.flat() as ChatMessageMention[];
    // eslint-disable-next-line @typescript-eslint/non-nullable-type-assertion-style
    const teamsMention = flatMentions.find(
      m => m.id?.toString() === mention?.id && m.mentionText === mention?.displayText
    ) as ChatMessageMention;

    const user = teamsMention?.mentioned?.user as User;
    if (user) {
      const MGTMention = (_props: MgtTemplateProps) => {
        return defaultRenderer(mention);
      };
      render = (
        <Person userId={user?.id} personCardInteraction={PersonCardInteraction.hover}>
          <MGTMention template="default" />
        </Person>
      );

      //   render = <button>{mention.displayText}</button>;
    }
    return render;
  };
};
