import React, { useState } from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from './ChatItem';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const { value } = props.dataContext;
  const chats: Chat[] = value;
  const [selectedChat, setSelectedChat] = useState<GraphChat>(props.selectedChat || chats[0]);

  React.useEffect(() => {
    props.onSelected(selectedChat);
  }, [selectedChat]);

  const onChatSelected = React.useCallback(
    (e: GraphChat) => {
      setSelectedChat(e);
    },
    [setSelectedChat]
  );

  const isChatActive = (chat: Chat) => {
    if (selectedChat) {
      return selectedChat && chat.id === selectedChat?.id;
    }

    return false;
  };

  console.log('chats', chats);
  return (
    <div>
      {chats.map((c, index) => (
        <ChatItem
          key={c.id}
          chat={c}
          isSelected={(!selectedChat && index === 0) || isChatActive(c)}
          onSelected={onChatSelected}
        />
      ))}
    </div>
  );
};

export default ChatListTemplate;
