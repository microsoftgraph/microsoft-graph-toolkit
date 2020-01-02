import addons, { makeDecorator } from '@storybook/addons';
import { Providers } from '../dist/es6/Providers';
import { ProviderState } from '../dist/es6/providers/IProvider';
import { MsalProvider } from '../dist/es6/providers/MsalProvider';
import { MockProvider } from '../dist/es6/mock/MockProvider';
import { CLIENTID, SETPROVIDER_EVENT, GETPROVIDER_EVENT } from './env';

export const withSignIn = makeDecorator({
  name: `withSignIn`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const mockProvider = new MockProvider(true);

    const channel = addons.getChannel();

    channel.on(SETPROVIDER_EVENT, params => {
      const currentProvider = Providers.globalProvider;
      if (params.state === ProviderState.SignedIn && (!currentProvider || currentProvider === mockProvider)) {
        Providers.globalProvider = new MsalProvider({
          clientId: CLIENTID
        });
      } else if (params.state !== ProviderState.SignedIn && currentProvider !== mockProvider) {
        Providers.globalProvider = mockProvider;
      }
    });

    // Our simple API above simply sets the notes parameter to a string,
    // which we send to the channel
    channel.emit(GETPROVIDER_EVENT, { type: 'getProvider' });
    // we can also add subscriptions here using channel.on('eventName', callback);

    return getStory(context);
  }
});
