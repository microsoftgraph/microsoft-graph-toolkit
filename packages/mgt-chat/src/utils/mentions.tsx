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
    const flatMentions = mentions.flat();
    const teamsMention: ChatMessageMention | undefined = flatMentions.find(
      m => m.id?.toString() === mention?.id && m.mentionText === mention?.displayText
    );

    const user = teamsMention?.mentioned?.user as User;
    if (user) {
      const me = user?.id === chatState.userId;
      const MgtMeMention = (_props: MgtTemplateProps) => {
        return defaultRenderer(mention);
      };
      const MgtOtherMention = () => {
        return <p className="otherMention">{mention.displayText}</p>;
      };
      const MgtMention = me ? MgtMeMention : MgtOtherMention;
      render = (
        <Person userId={user?.id} personCardInteraction={PersonCardInteraction.hover}>
          <MgtMention />
        </Person>
      );
    }
    return render;
  };
};
