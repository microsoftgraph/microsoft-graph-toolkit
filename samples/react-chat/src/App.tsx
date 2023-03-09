import React, { memo, useCallback, useState } from 'react';
import './App.css';
import { Get, Login, Picker } from '@microsoft/mgt-react';
import { MgtChat } from '@microsoft/mgt-chat';
import { Chat } from '@microsoft/microsoft-graph-types';
import ChatListTemplate from './components/ChatListTemplate/ChatListTemplate';

function App() {
  const [chatId, setChatId] = useState<string>();
  const chatSelected = useCallback(
    (e: Chat) => {
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
        <Get resource="me/chats?$expand=members" scopes={['chat.read']}>
          <ChatListTemplate template="default" onSelected={chatSelected} />
        </Get>
        {chatId && <MgtChat chatId={chatId} />}
      </header>
    </div>
  );
}

export default App;
