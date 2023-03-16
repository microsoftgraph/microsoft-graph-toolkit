import React, { memo } from 'react';
import { Chat, ConversationMember } from '@microsoft/microsoft-graph-types';
import { styles } from './chat-header.styles';

interface ChatHeaderProps {
  chat?: Chat;
}

const reduceToFirstNamesList = (participants: ConversationMember[]) => {
  return participants
    .map(p => {
      if (p.displayName?.includes(' ')) {
        return p.displayName.split(' ')[0];
      }
      return p.displayName || p.id;
    })
    .join(', ');
};

const ChatHeader = memo(({ chat }: ChatHeaderProps) => {
  return <div className={styles.chatHeader}>{reduceToFirstNamesList(chat?.members || [])}</div>;
});

export default ChatHeader;
