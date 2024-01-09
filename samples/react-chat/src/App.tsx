import React, { memo, useCallback, useState } from 'react';
import './App.css';
import { Get, Login } from '@microsoft/mgt-react';
import { Chat, ChatList, NewChat, ChatListButtonItem, ChatListMenuItem } from '@microsoft/mgt-chat';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { ChatAdd24Filled, ChatAdd24Regular, bundleIcon } from '@fluentui/react-icons';

const ChatAddIconBundle = bundleIcon(ChatAdd24Filled, ChatAdd24Regular);

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

  return (
    <Get
      resource="me/chats?$expand=members,lastMessagePreview&orderby=lastMessagePreview/createdDateTime desc"
      scopes={['chat.read']}
      cacheEnabled={false}
    >
      <ChatList menuItems={menus} buttonItems={buttons} onSelected={onSelected} />
    </Get>
  );
});

function App() {
  // Copied from the sample ChatItem.tsx
  // This should probably move up to ChatList.tsx so that all the ChatListItems can share myId
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
        <div className="chat-selector">
          <ChatListWrapper onSelected={chatSelected} />
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
        <div className="chat-pane">{chatId && <Chat chatId={chatId} />}</div>
      </main>
    </div>
  );
}

export default App;
