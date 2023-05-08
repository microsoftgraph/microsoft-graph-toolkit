import { Providers } from '@microsoft/mgt-element';
import { Chat, AadUserConversationMember } from '@microsoft/microsoft-graph-types';
import React, { useCallback, useEffect, useState } from 'react';

export interface ChatInteractionProps {
  onSelected: (selected: Chat) => void;
}

interface ChatItemProps {
  chat: Chat;
}

const ChatItem = ({ chat, onSelected }: ChatItemProps & ChatInteractionProps) => {
  const [myId, setMyId] = useState<string>();

  useEffect(() => {
    const getMyId = async () => {
      const me = await Providers.me();
      setMyId(me.id);
    };
    if (!myId) {
      void getMyId();
    }
  }, [myId]);

  const inferTitle = useCallback(
    (chat: Chat) => {
      if (chat.chatType === 'oneOnOne' && chat.members) {
        const other = chat.members.find(m => (m as AadUserConversationMember).userId !== myId);
        return other
          ? `Chat with ${other?.displayName || (other as AadUserConversationMember)?.email || other?.id}`
          : 'Chat with myself';
      }
      return chat.topic || chat.chatType;
    },
    [myId]
  );

  return <li onClick={() => onSelected(chat)}>{inferTitle(chat)}</li>;
};

export default ChatItem;
