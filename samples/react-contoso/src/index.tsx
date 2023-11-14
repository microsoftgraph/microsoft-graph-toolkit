import ReactDOM from 'react-dom';
import { App } from './App';
import { mergeStyles } from '@fluentui/react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider/dist/es6/exports';
import { Providers, LoginType } from '@microsoft/mgt-element';
import { MgtPersonCardConfig } from '@microsoft/mgt-components/dist/es6/exports';

MgtPersonCardConfig.isSendMessageVisible = false;

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
    'Chat.Create',
    'Chat.Read',
    'Chat.ReadBasic',
    'Chat.ReadWrite',
    'ChatMember.ReadWrite',
    'ChatMessage.Send',
    'ExternalItem.Read.All',
    'Files.Read',
    'Files.Read.All',
    'Files.ReadWrite.All',
    'Group.Read.All',
    'Group.ReadWrite.All',
    'Mail.Read',
    'Mail.ReadBasic',
    'People.Read',
    'People.Read.All',
    'Presence.Read.All',
    'User.Read',
    'Sites.Read.All',
    'Sites.ReadWrite.All',
    'Tasks.Read',
    'Tasks.ReadWrite',
    'Team.ReadBasic.All',
    'User.ReadBasic.All',
    'User.Read.All'
  ]
});

ReactDOM.render(<App />, document.getElementById('root'));
