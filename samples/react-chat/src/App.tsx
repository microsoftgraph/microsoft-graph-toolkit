import React from 'react';
import './App.css';
import { Login } from '@microsoft/mgt-react';
import { MgtChat } from '@microsoft/mgt-chat';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        Mgt Chat test harness
        <br />
        <Login />
        <MgtChat chatId="123456789" />
      </header>
    </div>
  );
}

export default App;
