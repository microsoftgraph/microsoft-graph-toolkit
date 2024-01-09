import React, { useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles, Button, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { StatefulGraphChatClient } from '../../statefulClient/StatefulGraphChatClient';
import { useGraphChatClient } from '../../statefulClient/useGraphChatClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';

export interface IChatListItemInteractionProps {
  onSelected: (e: GraphChat) => void;
}

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  linkContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  loadMore: {
    textDecorationLine: 'none',
    fontSize: '1.2em',
    fontWeight: 'bold',
    '&:hover': {
      textDecorationLine: 'none' // This removes the underline when hovering
    }
  },
  button: {
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%'
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = (
  props: MgtTemplateProps &
    IChatListItemInteractionProps &
    IChatListMenuItemsProps & {
      buttonItems?: ChatListButtonItem[];
    }
) => {
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

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;
  const chatListMenuItems = props.menuItems === undefined ? [] : props.menuItems;
  chatListMenuItems.unshift({
    displayText: 'Mark all as read',
    onClick: () => {}
  });

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <div>
          <div className={styles.headerContainer}>
            <ChatListHeader buttonItems={chatListButtonItems} menuItems={chatListMenuItems} />
          </div>
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
          <div className={styles.linkContainer}>
            <Link href="#" className={styles.loadMore}>
              load more
            </Link>
          </div>
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
