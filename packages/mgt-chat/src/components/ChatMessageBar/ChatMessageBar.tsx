import React from 'react';

import { MessageBar, MessageBarType } from '@fluentui/react';

interface ChatMessageBarProps {
  messageBarType: MessageBarType;
  message: string;
}

const ChatMessageBar = ({ messageBarType, message }: ChatMessageBarProps) => {
  return <MessageBar messageBarType={messageBarType}> {message}</MessageBar>;
};

export default ChatMessageBar;
