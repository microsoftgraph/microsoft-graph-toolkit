import * as React from 'react';
import { PageHeader } from '../components/PageHeader';
import { Get } from '@microsoft/mgt-react';
import { Loading } from '../components/Loading';
import {
  shorthands,
  makeStyles,
  mergeClasses,
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle
} from '@fluentui/react-components';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { Chat, NewChat } from '@microsoft/mgt-chat';
import ChatListTemplate from './Chats/ChatListTemplate';

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

const getPreviousDate = (months: number) => {
  const date = new Date();
  date.setMonth(date.getMonth() - months);
  return date.toISOString();
};

const nextResourceUrl = () =>
  `me/chats?$expand=members,lastMessagePreview&$orderBy=lastMessagePreview/createdDateTime desc&$filter=viewpoint/lastMessageReadDateTime ge ${getPreviousDate(
    9
  )}`;

const ChatPage: React.FunctionComponent = () => {
  const styles = useStyles();

  const [resourceUrl, setResourceUrl] = React.useState(nextResourceUrl);

  const [selectedChat, setSelectedChat] = React.useState<GraphChat>();
  const [isNewChatOpen, setIsNewChatOpen] = React.useState(false);

  const onChatCreated = (e: GraphChat) => {
    if (e.id !== selectedChat?.id && isNewChatOpen) {
      setIsNewChatOpen(false);
      setResourceUrl(nextResourceUrl);
    }
    setSelectedChat(e);
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
              <DialogTrigger disableButtonEnhancement>
                <Button appearance="primary" onClick={() => setIsNewChatOpen(true)}>
                  New Chat
                </Button>
              </DialogTrigger>
              <DialogSurface className={styles.dialogSurface}>
                <DialogBody className={styles.dialog}>
                  <DialogTitle>New Chat</DialogTitle>
                  <NewChat
                    onChatCreated={onChatCreated}
                    onCancelClicked={() => {
                      setIsNewChatOpen(false);
                    }}
                  ></NewChat>
                </DialogBody>
              </DialogSurface>
            </Dialog>
          </div>
          <ChatList onChatSelected={setSelectedChat} resourceUrl={resourceUrl}></ChatList>
        </div>
        <div className={styles.side}>{selectedChat && <Chat chatId={selectedChat.id!}></Chat>}</div>
      </div>
    </>
  );
};

interface ChatListProps {
  onChatSelected: (e: GraphChat) => void;
  resourceUrl: string;
}

const ChatList = React.memo((props: ChatListProps) => {
  return (
    <Get resource={props.resourceUrl} scopes={['chat.read']}>
      <ChatListTemplate template="default" onChatSelected={props.onChatSelected}></ChatListTemplate>
      <Loading template="loading" message={'Loading your chats...'}></Loading>
    </Get>
  );
});

export default ChatPage;
