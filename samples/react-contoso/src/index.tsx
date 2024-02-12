import ReactDOM from 'react-dom';
import { Suspense } from 'react';
import { mergeStyles } from '@fluentui/react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { Providers, LoginType, customElementHelper } from '@microsoft/mgt-element';
import { lazy } from 'react';
const App = lazy(() => import('./App'));

customElementHelper.withDisambiguation('foo');

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
  clientId: import.meta.env.VITE_CLIENT_ID!,
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

ReactDOM.render(
  <Suspense fallback="...">
    <App />
  </Suspense>,
  document.getElementById('root')
);
