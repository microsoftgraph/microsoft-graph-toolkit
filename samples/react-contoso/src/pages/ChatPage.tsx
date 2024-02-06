import * as React from 'react';
import { memo, useCallback } from 'react';
import { PageHeader } from '../components/PageHeader';
import {
  shorthands,
  makeStyles,
  mergeClasses,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle
} from '@fluentui/react-components';
import { Chat as GraphChat, ChatMessage } from '@microsoft/microsoft-graph-types';
import { ChatList, Chat, NewChat, ChatListButtonItem, ChatListMenuItem } from '@microsoft/mgt-chat';
import { Compose24Filled, Compose24Regular, bundleIcon } from '@fluentui/react-icons';

const ChatAddIconBundle = bundleIcon(Compose24Filled, Compose24Regular);

export const ChatAddIcon = (): JSX.Element => {
  const iconColor = 'var(--colorBrandForeground2)';
  return <ChatAddIconBundle color={iconColor} />;
};

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'row'
  },
  panels: {
    ...shorthands.padding('10px')
  },
  main: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '300px',
    minWidth: '300px',
    ...shorthands.overflow('auto'),
    maxHeight: '80vh',
    borderRightColor: 'var(--neutral-stroke-rest)',
    borderRightStyle: 'solid',
    borderRightWidth: '1px'
  },
  side: {
    display: 'flex',
    flexDirection: 'column',
    flexWrap: 'nowrap',
    width: '80%',
    ...shorthands.overflow('auto'),
    maxHeight: '80vh',
    height: '100%'
  },
  newChat: {
    paddingBottom: '10px',
    marginRight: '0px',
    marginLeft: 'auto'
  },
  dialog: {
    display: 'block'
  },
  dialogSurface: {
    contain: 'unset'
  }
});

interface ChatListWrapperProps {
  onSelected: (e: GraphChat) => void;
  onNewChat: () => void;
  selectedChatId: string | undefined;
}

const ChatListWrapper = memo(({ onSelected, onNewChat, selectedChatId }: ChatListWrapperProps) => {
  const buttons: ChatListButtonItem[] = [
    {
      renderIcon: () => <ChatAddIcon />,
      onClick: onNewChat
    }
  ];
  const menus: ChatListMenuItem[] = [
    {
      displayText: 'My custom menu item',
      onClick: () => console.log('My custom menu item clicked')
    }
  ];
  const onAllMessagesRead = useCallback((chatIds: string[]) => {
    console.log(`Number of chats marked as read: ${chatIds.length}`);
  }, []);
  const onLoaded = useCallback(() => {
    console.log('Chat threads loaded.');
  }, []);
  const onMessageReceived = useCallback((msg: ChatMessage) => {
    console.log('SampleChatLog: Message received', msg);
  }, []);

  return (
    <ChatList
      selectedChatId={selectedChatId}
      onLoaded={onLoaded}
      chatThreadsPerPage={10}
      menuItems={menus}
      buttonItems={buttons}
      onSelected={onSelected}
      onMessageReceived={onMessageReceived}
      onAllMessagesRead={onAllMessagesRead}
    />
  );
});

const ChatPage: React.FunctionComponent = () => {
  const styles = useStyles();
  const [chatId, setChatId] = React.useState<string>('');
  const [isNewChatOpen, setIsNewChatOpen] = React.useState(false);

  const onChatSelected = React.useCallback((e: GraphChat) => {
    if (chatId !== e.id) {
      setChatId(e.id ?? '');
    }
  }, []);

  const onNewChat = React.useCallback(() => {
    setIsNewChatOpen(true);
  }, []);

  const onChatCreated = (e: GraphChat) => {
    setIsNewChatOpen(false);
    if (chatId !== e.id) {
      setChatId(e.id ?? '');
    }
  };

  return (
    <>
      <PageHeader
        title={'Chats'}
        description={'Stay in touch with your teammates and navigate your chats'}
      ></PageHeader>

      <div className={styles.container}>
        <div className={mergeClasses(styles.panels, styles.main)}>
          <div className={styles.newChat}>
            <Dialog open={isNewChatOpen}>
              <DialogSurface className={styles.dialogSurface}>
                <DialogBody className={styles.dialog}>
                  <DialogTitle>New Chat</DialogTitle>
                  <NewChat onChatCreated={onChatCreated} onCancelClicked={() => setIsNewChatOpen(false)}></NewChat>
                </DialogBody>
              </DialogSurface>
            </Dialog>
          </div>
          <ChatListWrapper selectedChatId={chatId} onSelected={onChatSelected} onNewChat={onNewChat} />
        </div>
        <div className={styles.side}>
          <Chat chatId={chatId} />
        </div>
      </div>
    </>
  );
};

export default ChatPage;
