import ReactDOM from 'react-dom';
import { App } from './App';
import { mergeStyles } from '@fluentui/react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider/dist/es6/exports';
import { Providers, LoginType } from '@microsoft/mgt-element';

// Inject some global styles
mergeStyles({
  ':global(body,html,#root)': {
    margin: 0,
    padding: 0,
    height: '100vh',
    overflow: 'hidden'
  }
});

Providers.globalProvider = new Msal2Provider({
  clientId: process.env.REACT_APP_CLIENT_ID!,
  loginType: LoginType.Redirect,

  scopes: [
    'Bookmark.Read.All',
    'Calendars.Read',
    'Channel.ReadBasic.All',
    'ExternalItem.Read.All',
    'Files.ReadWrite.All',
    'Group.ReadWrite.All',
    'Mail.Read',
    'People.Read',
    'Presence.Read.All',
    'User.Read',
    'Sites.ReadWrite.All',
    'Tasks.ReadWrite',
    'Team.ReadBasic.All',
    'TermStore.Read.All',
    'User.Read.All'
  ]
});

ReactDOM.render(<App />, document.getElementById('root'));
