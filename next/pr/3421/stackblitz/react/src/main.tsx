import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import { Providers, MockProvider } from '@microsoft/mgt-element';
import { ThemeToggle } from '@microsoft/mgt-react';
import './main.css';

Providers.globalProvider = new MockProvider(true);

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <header>
      <ThemeToggle darkModeActive={false}  />
    </header>
    <App />
  </React.StrictMode>,
)
