import React, { useCallback, useState } from 'react';
import './App.css';
import { Get, Login } from '@microsoft/mgt-react';
import { Chat, NewChat } from '@microsoft/mgt-chat';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import ChatListTemplate from './components/ChatListTemplate/ChatListTemplate';

function App() {
  const [chatId, setChatId] = useState<string>();
  const chatSelected = useCallback((e: GraphChat) => {
    setChatId(e.id);
  }, []);

  const [showNewChat, setShowNewChat] = useState<boolean>(false);
  const onChatCreated = useCallback((chat: GraphChat) => {
    console.log('chat created', chat);
    setChatId(chat.id);
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
          <Get resource="me/chats?$expand=members" scopes={['chat.read']} cacheEnabled={true}>
            <ChatListTemplate template="default" onSelected={chatSelected} />
          </Get>
          Selected chat: {chatId}
          <br />
          <button onClick={() => setShowNewChat(true)}>New Chat</button>
          {showNewChat && (
            <div className="new-chat">
              <NewChat onChatCreated={onChatCreated} onCancelClicked={() => setShowNewChat(false)} />
            </div>
          )}
        </div>
        <div className="chat-pane">{chatId && <Chat chatId={chatId} />}</div>
      </main>
    </div>
  );
}

export default App;
