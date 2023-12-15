import React, { useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { SampleChats } from '../ChatListItem/sampleData';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';

// this is a stub to move the logic here that should end up here.
export const ChatList = (chatSelected: { chatSelected: (e: GraphChat) => void }) => {
  // TODO: change this to use StatefulGraphChatListClient
  const chatClient: StatefulGraphChatClient = useGraphChatClient('');
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [chatClient]);

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <ChatListItem chat={SampleChats.group.SampleGroupChat} myId={chatState.userId} onSelected={chatSelected} />
        <ChatListItem
          chat={SampleChats.group.SampleGroupChatMembershipChange}
          myId={chatState.userId}
          onSelected={chatSelected}
        />
        <ChatListItem chat={SampleChats.oneOnOne.SampleChat} myId={chatState.userId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleSelfChat} myId={chatState.userId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleTodayChat} myId={chatState.userId} onSelected={chatSelected} />
        <ChatListItem
          chat={SampleChats.oneOnOne.SampleYesterdayChat}
          myId={chatState.userId}
          onSelected={chatSelected}
        />
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
