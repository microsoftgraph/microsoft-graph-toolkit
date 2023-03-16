import React, { memo, useCallback, useState } from 'react';
import './App.css';
import { Get, Login } from '@microsoft/mgt-react';
import { Chat } from '@microsoft/mgt-chat';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import ChatListTemplate from './components/ChatListTemplate/ChatListTemplate';

function App() {
  const [chatId, setChatId] = useState<string>();
  const chatSelected = useCallback(
    (e: GraphChat) => {
      setChatId(e.id);
    },
    [setChatId]
  );

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
        </div>
        <div className="chat-pane">{chatId && <Chat chatId={chatId} />}</div>
      </main>
    </div>
  );
}

export default App;
