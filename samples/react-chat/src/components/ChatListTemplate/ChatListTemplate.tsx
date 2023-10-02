import React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from '../ChatItem/ChatItem';

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const { value } = props.dataContext as { value: Chat[] };
  const chats: Chat[] = value;
  // Select a default chat to display
  // props.onSelected(chats[0]);
  return (
    <ul>
      {chats.map(c => (
        <ChatItem key={c.id} chat={c} onSelected={props.onSelected} />
      ))}
    </ul>
  );
};

export default ChatListTemplate;
