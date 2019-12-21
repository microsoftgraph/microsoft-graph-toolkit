import addons, { makeDecorator } from '@storybook/addons';
import { Providers } from '../dist/es6/Providers';
import { ProviderState } from '../dist/es6/providers/IProvider';
import { MsalProvider } from '../dist/es6/providers/MsalProvider';
import { MockProvider } from '../dist/es6/mock/MockProvider';

export const withSignIn = makeDecorator({
  name: `withSignIn`,
  parameterName: 'myParameter',
  skipIfNoParametersOrOptions: false,
  wrapper: (getStory, context, { parameters }) => {
    const mockProvider = new MockProvider(true);

    const channel = addons.getChannel();

    channel.on('mgt/setProvider', params => {
      console.log('setProvider', params);
      const currentProvider = Providers.globalProvider;
      if (params.state === ProviderState.SignedIn && (!currentProvider || currentProvider === mockProvider)) {
        Providers.globalProvider = new MsalProvider({
          clientId: 'a974dfa0-9f57-49b9-95db-90f04ce2111a'
        });
        console.log('setting msal');
      } else if (params.state !== ProviderState.SignedIn && currentProvider !== mockProvider) {
        Providers.globalProvider = mockProvider;
        console.log('setting mock');
      }
    });

    // Our simple API above simply sets the notes parameter to a string,
    // which we send to the channel
    channel.emit('mgt/getProvider', { type: 'getProvider' });
    // we can also add subscriptions here using channel.on('eventName', callback);

    return getStory(context);
  }
});
