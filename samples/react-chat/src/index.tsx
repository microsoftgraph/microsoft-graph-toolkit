import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import reportWebVitals from './reportWebVitals';
import { Providers } from '@microsoft/mgt-react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { allChatListScopes, brokerSettings, GraphConfig } from '@microsoft/mgt-chat';

brokerSettings.defaultSubscriptionLifetimeInMinutes = 7;
brokerSettings.renewalThreshold = 65;
brokerSettings.renewalTimerInterval = 15;

Providers.globalProvider = new Msal2Provider({
  baseURL: GraphConfig.graphEndpoint,
  clientId: 'ed072e38-e76e-45ae-ab76-073cb95495bb',
  scopes: allChatListScopes
});

ReactDOM.render(<App />, document.getElementById('root'));

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals(console.debug);
