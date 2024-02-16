import { useCallback, useState } from 'react';
import './App.css';
import { Login } from '@microsoft/mgt-react';
import { Chat, ChatList, NewChat, ChatListButtonItem, ChatListMenuItem, IChatListActions } from '@microsoft/mgt-chat';
import { ChatMessage, Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { Compose24Filled, Compose24Regular, bundleIcon } from '@fluentui/react-icons';
import { GraphChatThread } from '../../../packages/mgt-chat/src/statefulClient/StatefulGraphChatListClient';

const ChatAddIconBundle = bundleIcon(Compose24Filled, Compose24Regular);

export const ChatAddIcon = (): JSX.Element => {
  const iconColor = 'var(--colorBrandForeground2)';
  return <ChatAddIconBundle color={iconColor} />;
};

function App() {
  const [chatId, setChatId] = useState<string>('');
  const [showNewChat, setShowNewChat] = useState<boolean>(false);

  const onChatSelected = useCallback((e: GraphChatThread) => {
    console.log('Selected: ', e.id);
    setChatId(e.id ?? '');
  }, []);

  const onChatCreated = useCallback((chat: GraphChat) => {
    setChatId(chat.id ?? '');
    setShowNewChat(false);
  }, []);

  const buttons: ChatListButtonItem[] = [
    {
      renderIcon: () => <ChatAddIcon />,
      onClick: () => setShowNewChat(true)
    }
  ];

  const menus: ChatListMenuItem[] = [
    {
      displayText: 'Mark all as read',
      onClick: (actions: IChatListActions) => actions.markAllChatThreadsAsRead()
    },
    {
      displayText: 'My custom menu item',
      onClick: () => console.log('My custom menu item clicked')
    }
  ];

  const onAllMessagesRead = useCallback((chatIds: string[]) => {
    console.log(`Number of chats marked as read: ${chatIds.length}`);
  }, []);

  const onLoaded = useCallback((chatThreads: GraphChatThread[]) => {
    console.log('Chat threads loaded: ', chatThreads.length);
  }, []);

  const onMessageReceived = useCallback((msg: ChatMessage) => {
    console.log('SampleChatLog: Message received', msg);
  }, []);

  const onConnectionChanged = useCallback((connected: boolean) => {
    console.log('Connection changed: ', connected);
  }, []);

  const onUnselected = useCallback((chatThread: GraphChatThread) => {
    console.log('Unselected: ', chatThread.id);
  }, []);

  return (
    <div className="App">
      <header className="App-header">
        Mgt Chat test harness
        <br />
        <Login />
      </header>
      <main className="main">
        <div className="chat-selector">
          <br />
          <button onClick={() => setChatId('')}>Clear selected chat</button>
          <br />
          <button onClick={() => setShowNewChat(true)}>New Chat</button>
          Selected chat: {chatId}
          <br />
          {showNewChat && (
            <div className="new-chat">
              <NewChat onChatCreated={onChatCreated} onCancelClicked={() => setShowNewChat(false)} mode="auto" />
            </div>
          )}
        </div>
        <div className="chatlist-pane">
          <ChatList
            onLoaded={onLoaded}
            chatThreadsPerPage={10}
            menuItems={menus}
            buttonItems={buttons}
            onSelected={onChatSelected}
            onMessageReceived={onMessageReceived}
            onAllMessagesRead={onAllMessagesRead}
            onConnectionChanged={onConnectionChanged}
            onUnselected={onUnselected}
          />
        </div>
        {/* NOTE: removed the chatId guard as this case has an error state. */}
        <div className="chat-pane">{<Chat chatId={chatId} />}</div>
      </main>
    </div>
  );
}

export default App;
