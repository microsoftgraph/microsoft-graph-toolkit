import React, { useState } from 'react';
import { MgtTemplateProps, Providers } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { ChatItem } from './ChatItem';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { List, ListItem } from '@fluentui/react-northstar';
import { makeStyles, shorthands } from '@fluentui/react-components';
import { iconFilledClassName, iconRegularClassName } from '@fluentui/react-icons';

const useStyles = makeStyles({
  listItem: {
    listStyleType: 'none',
    width: '100%',
    ':focus-visible': {
      [`& .${iconFilledClassName}`]: {
        display: 'inline',
        color: 'var(--colorNeutralForeground2BrandHover)'
      },
      [`& .${iconRegularClassName}`]: {
        display: 'none'
      }
    }
  },
  selected: {
    backgroundColor: 'var(--colorNeutralBackground1Selected)'
  },
  list: {
    fontWeight: 800,
    gridGap: '8px',
    ...shorthands.marginBlock('0'),
    ...shorthands.padding('0')
  }
});

export interface ChatInteractionProps {
  onChatSelected: (selected: Chat) => void;
}

const ChatListTemplate = (props: MgtTemplateProps & ChatInteractionProps) => {
  const styles = useStyles();
  const { value } = props.dataContext;
  const [chats] = useState<Chat[]>((value as Chat[]).filter(c => c.members?.length! > 1));
  const [userId, setUserId] = useState<string>();
  const [selectedChat, setSelectedChat] = useState<GraphChat | undefined>(chats.length > 0 ? chats[0] : undefined);

  const onChatSelected = (e: GraphChat) => {
    setSelectedChat(e);
  }

  // Set the selected chat to the first chat in the list
  // Fires only the first time the component is rendered
  React.useEffect(() => {
    if (selectedChat) {
      props.onChatSelected(selectedChat);
    }
  }, [props, selectedChat]);

  React.useEffect(() => {
    const getMyId = async () => {
      const me = await Providers.me();
      setUserId(me.id);
    };
    if (!userId) {
      void getMyId();
    }
  });

  const isChatActive = (chat: Chat) => {
    if (selectedChat) {
      return selectedChat && chat.id === selectedChat?.id;
    }

    return false;
  };

  return (
    <List navigable className={styles.list}>
      {chats.map((c, index) => (
        <ListItem
          key={c.id}
          index={index}
          className={styles.listItem}
          content={
            <div className={(!selectedChat && index === 0) || isChatActive(c) ? styles.selected : ''}>
              <ChatItem
                key={c.id}
                chat={c}
                userId={userId}
              />
            </div>
          }
          onClick={() => onChatSelected(c)}>
        </ListItem>
      ))}
    </List>
  );
};

export default ChatListTemplate;
