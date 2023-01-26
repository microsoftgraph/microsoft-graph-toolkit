
import React, { lazy, Suspense } from 'react';
import ReactDOM from 'react-dom';
import { customElementHelper, Providers, MockProvider } from '@microsoft/mgt-element';
// import { MsalProvider } from '@microsoft/mgt-msal-provider';
// import { Msal2Provider } from '@microsoft/mgt-msal2-provider';

customElementHelper.withDisambiguation('contoso');

// uncomment to use MSAL provider with your own ClientID (also comment out mock provider below)
// Providers.globalProvider = new MsalProvider({ clientId: 'a974dfa0-9f57-49b9-95db-90f04ce2111a' });
// Providers.globalProvider = new Msal2Provider({ clientId: '2dfea037-938a-4ed8-9b35-c05708a1b241' });

Providers.globalProvider = new MockProvider(true);

const App = lazy(() => import('./App'));
ReactDOM.render(<Suspense fallback='...'><App /></Suspense>, document.getElementById('root'));
