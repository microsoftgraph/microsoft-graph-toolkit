import React from 'react';

interface IMgtChatProps {
  chatId: string;
}

const MgtChat = ({ chatId }: IMgtChatProps) => {
  return <>Rendering chat: {chatId}</>;
};

export default MgtChat;
