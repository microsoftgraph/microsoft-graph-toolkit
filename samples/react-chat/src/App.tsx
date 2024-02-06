import React, { memo, useCallback, useState } from 'react';
import './App.css';
import { Login } from '@microsoft/mgt-react';
import { Chat, ChatList, NewChat, ChatListButtonItem, ChatListMenuItem } from '@microsoft/mgt-chat';
import { ChatMessage, Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { Compose24Filled, Compose24Regular, bundleIcon } from '@fluentui/react-icons';

const ChatAddIconBundle = bundleIcon(Compose24Filled, Compose24Regular);

export const ChatAddIcon = (): JSX.Element => {
  const iconColor = 'var(--colorBrandForeground2)';
  return <ChatAddIconBundle color={iconColor} />;
};

const ChatListWrapper = memo(({ onSelected }: { onSelected: (e: GraphChat) => void }) => {
  const buttons: ChatListButtonItem[] = [
    {
      renderIcon: () => <ChatAddIcon />,
      onClick: () => console.log('Add chat clicked')
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
  const onLoaded = () => {
    console.log('Chat threads loaded.');
  };
  const onMessageReceived = (msg: ChatMessage) => {
    console.log('SampleChatLog: Message received', msg);
  };

  return (
    <ChatList
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

function App() {
  const [chatId, setChatId] = useState<string>('');
  const [showNewChat, setShowNewChat] = useState<boolean>(false);

  const chatSelected = useCallback((e: GraphChat) => {
    setChatId(e.id ?? '');
  }, []);

  const onChatCreated = useCallback((chat: GraphChat) => {
    setChatId(chat.id ?? '');
    setShowNewChat(false);
  }, []);

  return (
    <div className="App">
      <header className="App-header">
        Mgt Chat test harness
        <br />
        <Login />
      </header>
      <main className="main">
        {/* NOTE: removed the chatId guard as this case has an error state. */}
        <div className="chat-pane">{<Chat chatId={chatId} />}</div>
      </main>
    </div>
  );
}

export default App;
