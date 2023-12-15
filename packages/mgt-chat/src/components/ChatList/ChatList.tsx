import React, { useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { SampleChats } from '../ChatListItem/sampleData';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { ProviderState, Providers } from '@microsoft/mgt-element';

// this is a stub to move the logic here that should end up here.
export const ChatList = (chatSelected: { chatSelected: (e: GraphChat) => void }) => {
  const [myId, setMyId] = useState<string>();

  useEffect(() => {
    const currentUserId = () => getCurrentUser()?.id.split('.')[0] || '';

    const getMyId = async () => {
      const id = currentUserId();
      // check if me is null
      if (!id) {
        console.error('Could not get current user id.');
      } else {
        setMyId(id);
      }
    };
    if (!myId) {
      getMyId();
    }
  }, [myId]);

  const getCurrentUser = () =>
    Providers.globalProvider.state === ProviderState.SignedIn
      ? Providers.globalProvider.getActiveAccount?.()
      : undefined;

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <ChatListItem chat={SampleChats.group.SampleGroupChat} myId={myId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.group.SampleGroupChatMembershipChange} myId={myId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleChat} myId={myId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleSelfChat} myId={myId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleTodayChat} myId={myId} onSelected={chatSelected} />
        <ChatListItem chat={SampleChats.oneOnOne.SampleYesterdayChat} myId={myId} onSelected={chatSelected} />
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;