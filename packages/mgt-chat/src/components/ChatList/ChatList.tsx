import React, { useEffect, useState } from 'react';
import { ChatListItem, IChatListItemInteractionProps } from '../ChatListItem/ChatListItem';
import { SampleChats } from '../ChatListItem/sampleData';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { makeStyles, shorthands } from '@fluentui/react-components';

const useStyles = makeStyles({
  chatList: {
    display: 'flex',
    flexDirection: 'column',
    maxWidth: '300px',
    width: '100%',
    height: '100%',
    ...shorthands.overflow('hidden')
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = (props: MgtTemplateProps & IChatListItemInteractionProps) => {
  // TODO: change this to use StatefulGraphChatListClient
  const chatClient: StatefulGraphChatClient = useGraphChatClient('');
  const [chatState, setChatState] = useState(chatClient.getState());
  useEffect(() => {
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offStateChange(setChatState);
    };
  }, [chatClient]);

  const { value } = props.dataContext as { value: GraphChat[] };
  const chats: GraphChat[] = value;

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        {chats.map(c => (
          <ChatListItem key={c.id} chat={c} myId={chatState.userId} onSelected={props.onSelected} />
        ))}
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
