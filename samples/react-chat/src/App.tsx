import React from 'react';
import './App.css';
import { MgtChat } from '@microsoft/mgt-chat';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        Mgt Chat test harness
        <br />
        <MgtChat chatId="123456789" />
      </header>
    </div>
  );
}

export default App;
