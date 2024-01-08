import React, { useEffect, useState, useCallback } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { SampleChats } from '../ChatListItem/sampleData';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { FluentProvider, webLightTheme, Button, makeStyles, shorthands } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';

export interface IChatListItemInteractionProps {
  onSelected: (e: GraphChat) => void;
}

const useStyles = makeStyles({
  button: {
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%'
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = (props: MgtTemplateProps & IChatListItemInteractionProps) => {
  const styles = useStyles();

  // TODO: change this to use StatefulGraphChatListClient
  const chatClient: StatefulGraphChatClient = useGraphChatClient('');
  const [chatState, setChatState] = useState(chatClient.getState());
  const [selectedItem, setSelectedItem] = useState<string>();

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
          <Button
            className={styles.button}
            key={c.id}
            onClick={() => {
              // set selected state only once per click event
              if (c.id !== selectedItem) {
                setSelectedItem(c.id);
                props.onSelected(c);
              }
            }}
          >
            <ChatListItem
              key={c.id}
              chat={c}
              myId={chatState.userId}
              isSelected={c.id === selectedItem}
              isRead={false}
            />
          </Button>
        ))}
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
