import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';

import { Providers, MockProvider } from '@microsoft/mgt-element';

// uncomment to use MSAL provider with your own ClientID (also comment out mock provider below)
// import { MsalProvider } from '@microsoft/mgt-msal-provider';
// Providers.globalProvider = new MsalProvider({ clientId: 'a974dfa0-9f57-49b9-95db-90f04ce2111a' });

Providers.globalProvider = new MockProvider(true);

ReactDOM.render(<App />, document.getElementById('root'));
