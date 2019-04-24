# Providers

The Microsoft Graph Toolkit defines Providers that enable all components to be aware of the current authentication state and initiate calls to the Microsoft Graph without the developer having to write the code themselves. Each provider provides implementation for acquiring the necessary access token for calling the Microsoft Graph APIs.

In order for the components to use a provider, the `Providers.globalProvider` property must be set to a Provider you'd like to use.

Example of using the MsalProvider:

```js
Providers.globalProvider = new MsalProvider({
  clientId: "[CLIENT_ID]"
});
```

The toolkit implements several providers:

- [MsalProvider](./providers/msal.md)
- [SharePointProvider](./providers/sharepoint.md)
- [TeamsProvider](./providers/teams.md)
- Office Add-ins provider (coming soon)

## Get started

You can create a provider at any time, but it's recommended to create it before using any of the components.

### API

1. The `Providers` global variable exposes the following properties and functions

   - `globalProvider : IProvider`

     set this property to a provider you want to use globally. All components use this property to get a reference to the provider. Setting this property will fire the onProvidersChanged event

   - `function onProviderUpdated(callbackFunction)`

     the `callbackFunction` function will be called when a provider is changed or when the state of a Provider changes. A `ProvidersChangedState` enum value will be passed to the function to indicate what updated

## Implement your own provider

You can extend the `IProvider` abstract class to create your own provider.

### State

A provider must keep track of the authentication state and update the components when the state changes. The `IProvider` class already implements the `onStateChanged(eventHandler)` handler and the `state: ProviderState` property. You as a developer just need to use the `setState(state:ProviderState)` method in your implementation to update the state when it changes. Updating the state will fire the stateChanged event and update all the components automatically.

### Login/Logout

If your provider provides login or logout functionality, implement the `login(): Promise<void>` and `logout(): Promise<void>` methods. These methods are optional.

### Access Token

You must implement the `getAccessToken({'scopes': scopes}) : Promise<string>` method (`scopes` is a string array of scopes). This method is used to get a valid token before every call to the Microsoft Graph.

### Graph

The components use the Microsoft Graph Javascript SDK for all calls to the Microsoft Graph. Your provider must make the sdk available through the `graph` property. In you constructor, create a new Graph instance through

```js
this.graph = new Graph(this);
```

The `Graph` class is a light wrapper on top of the Microsoft Graph sdk.

## Making your own calls to the Microsoft Graph

All components can access the Microsoft Graph out of the box as long as the developer has initialized a provider (as described in the above section). To get a reference to the same Microsoft Graph SDK used by the components, first get a reference to the global IProvider and then use the `Graph` object:

```js
import { Providers } from "microsoft-graph-toolkit";

let provider = Providers.globalProvider;
if (provider) {
  let graphClient = provider.graph.client;
  let userDetails = await graphClient.api("me").get();
}
```

There might be cases were you will need to pass additional scopes depending on the api you are calling:

```js
import { prepScopes } from "microsoft-graph-toolkit";

graphClient
  .api("me")
  .middlewareOptions(prepScopes("user.read", "calendar.read"))
  .get();
```

The graph object is an instance of the [Microsoft Graph Javascript SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript) and you can use it to make any calls to the Microsoft Graph.
