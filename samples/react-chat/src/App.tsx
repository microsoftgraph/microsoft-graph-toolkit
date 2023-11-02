import React, { memo, useCallback, useState } from 'react';
import './App.css';
import { Get, Login } from '@microsoft/mgt-react';
import { Chat, NewChat } from '@microsoft/mgt-chat';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import ChatListTemplate from './components/ChatListTemplate/ChatListTemplate';

const ChatList = memo(({ chatSelected }: { chatSelected: (e: GraphChat) => void }) => {
  return (
    <Get resource="me/chats?$expand=members" scopes={['chat.read']} cacheEnabled={false}>
      <ChatListTemplate template="default" onSelected={chatSelected} />
    </Get>
  );
});

function App() {
  const [chatId, setChatId] = useState<string>('');
  const chatSelected = useCallback((e: GraphChat) => {
    setChatId(e.id ?? '');
  }, []);

  const [showNewChat, setShowNewChat] = useState<boolean>(false);
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
          <ChatList chatSelected={chatSelected} />
          Selected chat: {chatId}
          <br />
          <button onClick={() => setChatId('')}>Clear selected chat</button>
          <br />
          <button onClick={() => setShowNewChat(true)}>New Chat</button>
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
