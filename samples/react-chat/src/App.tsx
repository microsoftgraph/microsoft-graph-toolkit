import React, { memo, useCallback, useState, useEffect } from 'react';
import './App.css';
import { Get, Login } from '@microsoft/mgt-react';
import { Chat, NewChat } from '@microsoft/mgt-chat';
import { CacheService, log } from '@microsoft/mgt-element';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import ChatListTemplate from './components/ChatListTemplate/ChatListTemplate';

const ChatList = memo(({ chatSelected }: { chatSelected: (e: GraphChat) => void }) => {
  return (
    <Get resource="me/chats?$expand=members" scopes={['Chat.ReadWrite']} cacheEnabled={false}>
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

  // start a timer to force a re-render of presence
  const refreshIntervalForPresence = 2 * 60 * 1000;
  const [iteration, setIteration] = useState(0);
  useEffect(() => {
    // setup caching
    CacheService.config.users.isEnabled = true;
    CacheService.config.photos.isEnabled = true;
    CacheService.config.presence.isEnabled = true;

    // exit if no refresh interval for presence
    if (!refreshIntervalForPresence) return;

    // set the invalidation period to half the refresh interval
    CacheService.config.presence.invalidationPeriod = refreshIntervalForPresence / 2;

    // start refresh timer for presence
    log(`starting refresh timer for presence every ${refreshIntervalForPresence}ms`);
    const intervalId = setInterval(() => {
      setIteration(prevIteration => prevIteration + 1);
    }, refreshIntervalForPresence);

    // clear interval on unmount
    return () => clearInterval(intervalId);
  }, [refreshIntervalForPresence]);

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

        <div className="chat-pane">{chatId && <Chat chatId={chatId} iteration={iteration} />}</div>
      </main>
    </div>
  );
}

export default App;
