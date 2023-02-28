import React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from '../ChatItem/ChatItem';

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const { value } = props.dataContext;
  const chats: Chat[] = value;
  return (
    <ul>
      {chats.map(c => (
        <ChatItem chat={c} onSelected={props.onSelected} />
      ))}
    </ul>
  );
};

export default ChatListTemplate;
