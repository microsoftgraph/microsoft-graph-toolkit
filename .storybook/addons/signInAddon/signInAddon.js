import addons, { makeDecorator } from '@storybook/addons';
import { Providers } from '../../../packages/mgt-element/dist/es6/providers/Providers';
import { ProviderState } from '../../../packages/mgt-element/dist/es6/providers/IProvider';
import { Msal2Provider } from '../../../packages/providers/mgt-msal2-provider/dist/es6/Msal2Provider';
import { MockProvider } from '../../../packages/mgt-element/dist/es6/mock/MockProvider';
import { CLIENTID, SETPROVIDER_EVENT, GETPROVIDER_EVENT } from '../../env';

const _allow_signin = true;

export const withSignIn = makeDecorator({
  name: `withSignIn`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const mockProvider = new MockProvider(true);

    const channel = addons.getChannel();

    channel.on(SETPROVIDER_EVENT, params => {
      if (_allow_signin) {
        const currentProvider = Providers.globalProvider;
        if (params.state === ProviderState.SignedIn && (!currentProvider || currentProvider === mockProvider)) {
          Providers.globalProvider = new Msal2Provider({
            clientId: CLIENTID,
            scopes: [
              'user.read',
              'user.read.all',
              'mail.readBasic',
              'people.read',
              'people.read.all',
              'sites.read.all',
              'user.readbasic.all',
              'contacts.read',
              'presence.read',
              'presence.read.all',
              'tasks.readwrite',
              'tasks.read',
              'calendars.read',
              'group.read.all',
              'files.read',
              'files.read.all',
              'files.readwrite',
              'files.readwrite.all'
            ]
          });
        } else if (params.state !== ProviderState.SignedIn && currentProvider !== mockProvider) {
          Providers.globalProvider = mockProvider;
        }
      }
    });

    // Our simple API above simply sets the notes parameter to a string,
    // which we send to the channel
    channel.emit(GETPROVIDER_EVENT, { type: 'getProvider' });
    // we can also add subscriptions here using channel.on('eventName', callback);

    return getStory(context);
  }
});
