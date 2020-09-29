import 'react-app-polyfill/ie11';
import 'react-app-polyfill/stable';
// in theory, we should be able to add this to index.html
//  <script>
//    window.WebComponents = window.WebComponents || {};
//    window.WebComponents.root = '../node_modules/@webcomponents/webcomponentsjs/';
//  </script>
// and then
// import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-loader.js';
// but for some reason it doesn't on edge
// because of this, we'll just load the whole bundle.
import '../node_modules/@webcomponents/webcomponentsjs/webcomponents-bundle.js';
import React from 'react';
import ReactDOM from 'react-dom';
import App from './App';
import { Providers, MsalProvider } from '@microsoft/mgt';
import { MockProvider } from '@microsoft/mgt/dist/es6/mock/MockProvider';

Providers.globalProvider = new MockProvider(true);

// uncomment to use MSAL provider with your own ClientID (also comment above line)
// Providers.globalProvider = new MsalProvider({clientId: '{YOUR_CLIENT_ID}'});

ReactDOM.render(<App />, document.getElementById('root'));
