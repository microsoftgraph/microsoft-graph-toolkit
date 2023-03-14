import ReactDOM from 'react-dom';
import { App } from './App';
import { mergeStyles } from '@fluentui/react';
import { initializeIcons } from '@uifabric/icons';
import { ThemeProvider } from '@fluentui/react-theme-provider';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { Providers, LoginType } from '@microsoft/mgt-element';

initializeIcons(undefined, { disableWarnings: true });

// Inject some global styles
mergeStyles({
  ':global(body,html,#root)': {
    margin: 0,
    padding: 0,
    height: '100vh'
  }
});

Providers.globalProvider = new Msal2Provider({
  clientId: process.env.REACT_APP_CLIENT_ID!,
  loginType: LoginType.Popup,
  scopes: ['User.Read', 'Bookmark.Read.All', 'ExternalItem.Read.All']
});

ReactDOM.render(
  <ThemeProvider>
    <App />
  </ThemeProvider>,
  document.getElementById('root')
);
