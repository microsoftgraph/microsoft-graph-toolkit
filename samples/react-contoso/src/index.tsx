import ReactDOM from 'react-dom';
import { App } from './App';
import { mergeStyles } from '@fluentui/react';
import { initializeIcons } from '@uifabric/icons';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { Providers, LoginType } from '@microsoft/mgt-element';
import { allDashboardScopes } from './pages/DashboardPage';
import { allIncidentScopes } from './pages/Incident/Incident';

initializeIcons(undefined, { disableWarnings: true });

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
  scopes: ['User.Read', ...allDashboardScopes, ...allIncidentScopes]
});

ReactDOM.render(<App />, document.getElementById('root'));
