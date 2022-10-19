# Microsoft Graph Toolkit Base package

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-element?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-element)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-element` package contains all base classes that enable the providers and components to work together. Use this package to set the global provider, or to create your own providers and/or components that work with Microsoft Graph.

[See docs for full documentation](https://aka.ms/mgt-docs)

## Set and use the global provider

The `@microsoft/mgt-element` package exposes the `Providers` namespace that enables global usage of the authentication providers across your entire app. 

This example illustrates how to instantiate a new provider (MsalProvider in this case) and use it across your app:

1. Install the packages 

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-msal-provider
    ```

1. Create the provider

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {MsalProvider} from '@microsoft/mgt-msal-provider';

    // initialize the auth provider globally
    Providers.globalProvider = new MsalProvider({clientId: 'clientId'});
    ```

1. Use the provider to sign in and call the graph:

    ```ts
    import {Providers, ProviderState} from '@microsoft/mgt-element';

    const handleLoginClicked = async () => {
      await Providers.globalProvider.login();

      if (Providers.globalProvider.state === ProviderState.SignedIn) {
        let me = await Provider.globalProvider.graph.client.api('/me').get();
      }
    }
    ```

You can learn more about how to use the providers in the [documentation](https://learn.microsoft.com/graph/toolkit/providers).

The providers work well with the `@microsoft/mgt-components` package and all components use the provider automatically when they need to call Microsoft Graph.

## Create your own provider

In scenarios where you want to use the Providers namespace and/or add Microsoft Graph Toolkit components to an application with pre-existing authentication code, you can create a custom provider that hooks into your authentication mechanism. `@microsoft/mgt-element` enables two ways to create new providers:

### Create a Simple Provider

If you already have a function that returns `accessTokens`, you can use a SimpleProvider to wrap the function:

```ts
import {Providers, SimpleProvider} from '@microsoft/mgt-element';

function getAccessToken(scopes: string[]) {
  // return a promise with accessToken string
}

function login() {
  // login code - optional

  // make sure to set the state when signed in
  Providers.globalProvider.setState(ProviderState.SignedIn)
}

function logout() {
  // logout code - optional

  // make sure to set the state when signed out
  Providers.globalProvider.setState(ProviderState.SignedOut)
}

Provider.globalProvider = new SimpleProvider(getAccessToken, login, logout);
```

### Extend an IProvider

You can extend the IProvider abstract class to create your own provider. The IProvider is similar to the SimpleProvider in that it requires the developer to implement the `getAccessToken()` function.


See the [custom provider documentation](https://learn.microsoft.com/graph/toolkit/providers/custom) for more details on both ways to create custom providers.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
