import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';
import {Providers, MsalProvider} from '@microsoft/mgt';
import {MockProvider} from '@microsoft/mgt/dist/es6/mock/MockProvider';

Providers.globalProvider = new MockProvider();

// uncomment to use MSAL provider with your own ClientID (also comment above line)
// Providers.globalProvider = new MsalProvider({clientId: '{YOUR_CLIENT_ID}'});

ReactDOM.render(<App />, document.getElementById('root'));
