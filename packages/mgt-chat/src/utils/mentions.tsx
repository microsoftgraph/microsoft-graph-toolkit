import React from 'react';
import { MgtTemplateProps, Person } from '@microsoft/mgt-react';
import { ChatMessageMention, User } from '@microsoft/microsoft-graph-types';
import { GraphChatClient } from 'src/statefulClient/StatefulGraphChatClient';
import { Mention, MentionLookupOptions } from '@azure/communication-react';

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
        <Person className="mgt-person-mention" userId={user?.id} personCardInteraction="hover">
          <MgtMention />
        </Person>
      );
    }
    render = <>{render}&nbsp;</>;
    return render;
  };
};

export const mentionLookupOptionsWrapper = (chatState: GraphChatClient): MentionLookupOptions => {
  const participants = chatState.participants ?? [];
  const matchedResults: Record<string, string> = {};

  return {
    onQueryUpdated: (query: string): Promise<Mention[]> => {
      const results = participants.filter(p => p.displayName?.toLowerCase()?.includes(query.toLowerCase())) ?? [];
      const mentions: Mention[] = [];
      results.forEach((user, id) => {
        const idStr = `${id}`;
        mentions.push({ displayText: user?.displayName ?? '', id: idStr });
        matchedResults[idStr] = user?.userId ?? '';
      });
      return Promise.resolve(mentions);
    },
    onRenderSuggestionItem: (suggestion: Mention, onSuggestionSelected: (suggestion: Mention) => void): JSX.Element => {
      const userId = matchedResults[suggestion.id] ?? '';
      const key = userId ?? `${participants.length + 1}`;
      return (
        <Person key={key} userId={userId} view="oneline" onClick={() => onSuggestionSelected(suggestion)}></Person>
      );
    }
  };
};
