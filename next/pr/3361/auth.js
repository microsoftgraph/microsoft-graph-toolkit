import { Providers } from '../packages/mgt-element/dist/es6/providers/Providers';
import { Msal2Provider } from '../packages/providers/mgt-msal2-provider/dist/es6/Msal2Provider';
import { CLIENTID } from './env';

Providers.globalProvider = new Msal2Provider({
  clientId: CLIENTID
});
