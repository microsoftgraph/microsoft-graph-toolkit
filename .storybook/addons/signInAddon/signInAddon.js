import addons, { makeDecorator } from '@storybook/addons';
import { Providers } from '../../../packages/mgt-element/dist/providers/Providers';
import { ProviderState } from '../../../packages/mgt-element/dist/providers/IProvider';
import { MsalProvider } from '../../../packages/mgt/dist/es6/providers/MsalProvider';
import { MockProvider } from '../../../packages/mgt/dist/es6/mock/MockProvider';
import { CLIENTID, SETPROVIDER_EVENT, GETPROVIDER_EVENT } from '../../env';

const _allow_signin = false;

export const withSignIn = makeDecorator({
  name: `withSignIn`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const mockProvider = new MockProvider(true);

    Providers.globalProvider = mockProvider;

    const channel = addons.getChannel();

    channel.on(SETPROVIDER_EVENT, params => {
      if (_allow_signin) {
        const currentProvider = Providers.globalProvider;
        if (params.state === ProviderState.SignedIn && (!currentProvider || currentProvider === mockProvider)) {
          Providers.globalProvider = new MsalProvider({
            clientId: CLIENTID
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
