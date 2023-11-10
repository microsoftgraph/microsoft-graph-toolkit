import React, { useState } from 'react';
import { MgtTemplateProps, Providers } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatItem, { ChatInteractionProps } from './ChatItem';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const { value } = props.dataContext;
  const chats: Chat[] = value;
  const [userId, setUserId] = useState<string>();
  const [selectedChat, setSelectedChat] = useState<GraphChat>(/*props.selectedChat || */chats[0]);

  const onChatSelected = React.useCallback(
    (e: GraphChat) => {
      setSelectedChat(e);
      props.onSelected(selectedChat);
    },
    [setSelectedChat, selectedChat, props]
  );

  // Set the selected chat to the first chat in the list
  // Fires only the first time the component is rendered
  React.useEffect(() => {
    onChatSelected(selectedChat);
  });

  React.useEffect(() => {
    const getMyId = async () => {
      const me = await Providers.me();
      setUserId(me.id);
    };
    if (!userId) {
      void getMyId();
    }
  }, [userId]);

  const isChatActive = (chat: Chat) => {
    if (selectedChat) {
      return selectedChat && chat.id === selectedChat?.id;
    }

    return false;
  };

  return (
    <div>
      {chats.filter(c => c.members?.length! > 1).map((c, index) => (
        <ChatItem
          key={c.id}
          chat={c}
          isSelected={(!selectedChat && index === 0) || isChatActive(c)}
          onSelected={onChatSelected}
          userId={userId}
        />
      ))}
    </div>
  );
};

export default ChatListTemplate;
