import { useState } from 'react';
import { StatefulGraphChatClient } from './StatefulGraphChatClient';

export const useGraphChatClient = (chatId: string): StatefulGraphChatClient => {
  const [chatClient] = useState<StatefulGraphChatClient>(new StatefulGraphChatClient());
  chatClient.chatId = chatId;

  return chatClient;
};
