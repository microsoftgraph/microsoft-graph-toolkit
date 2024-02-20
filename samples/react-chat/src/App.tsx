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
  let sessionChatId = sessionStorage.getItem('chatId') ?? '';
  console.log('sessionChatId: ', sessionChatId);

  const [chatId, setChatId] = useState<string>(sessionChatId);
  const [chatThreadsPerPage, setChatThreadsPerPage] = useState<number>(10);
  const [showNewChat, setShowNewChat] = useState<boolean>(false);
  // we are using a different state to track the selected chat id fired from chat list.
  const [selectedChatListChatId, setSelectedChatListChatId] = useState<string>('');

  sessionStorage.clear();

  const saveChatAndRefresh = () => {
    if (chatId !== '') {
      console.log('setting chatId: ', chatId);
      sessionStorage.setItem('chatId', chatId);
      // force a page refesh, this will test setting the initial chat id.
      window.location.reload();
    }
  };

  const clearSelectedChat = () => {
    setChatId('');
    setSelectedChatListChatId('');
  };

  const onChatSelected = useCallback((e: GraphChatThread) => {
    console.log('Selected: ', e.id);
    setChatId(e.id ?? '');
    setSelectedChatListChatId(e.id ?? '');
  }, []);

  const onChatCreated = useCallback((chat: GraphChat) => {
    setChatId(chat.id ?? '');
    setSelectedChatListChatId(chat.id ?? '');
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

  const handleChange = useCallback((event: React.ChangeEvent<HTMLSelectElement>) => {
    setChatThreadsPerPage(parseInt(event.target.value));
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
          <button onClick={() => saveChatAndRefresh()}>Save selected chat and refresh</button>
          <br />
          <button onClick={() => clearSelectedChat()}>Clear selected chat</button>
          <br />
          <select value={chatThreadsPerPage} onChange={handleChange}>
            <option value="5">5</option>
            <option value="10">10</option>
            <option value="20">20</option>
            <option value="50">50</option>
            <option value="70">70</option>
          </select>
          <br />
          Chat threads per page: {chatThreadsPerPage}
          <br />
          <button onClick={() => setShowNewChat(true)}>New Chat</button>
          <br />
          Selected chat id: {chatId}
          <br />
          Selected chatlist chat id: {selectedChatListChatId}
          <br />
          {showNewChat && (
            <div className="new-chat">
              <NewChat onChatCreated={onChatCreated} onCancelClicked={() => setShowNewChat(false)} mode="auto" />
            </div>
          )}
        </div>
        <div className="chatlist-pane">
          <ChatList
            selectedChatId={chatId}
            onLoaded={onLoaded}
            chatThreadsPerPage={chatThreadsPerPage}
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
        <div className="chat-pane">{<Chat chatId={selectedChatListChatId} />}</div>
      </main>
    </div>
  );
}

export default App;
