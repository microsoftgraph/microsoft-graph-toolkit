# Custom provider

If you already have authentication code in your application, there are two ways to integrate it with the Providers in the toolkit.

- Create a new `SimpleProvider` by passing in a function for getting an access token, or
- Extend the `IProvider` abstract class

Let's take a look at each one in more details

## SimpleProvider

The SimpleProvider class can be instantiated by passing in a function that will return an accesstoken for passed in scopes.

```ts
let provider = new SimpleProvider((scopes: string[]) => {
  // return a promise with accessToken
});
```

In addition, you can also provide an optional `login` and `logout` functions that can handle the login and logout calls from the [Login](../components/login.md) component

```ts
function getAccessToken(scopes: string[]) {
  // return a promise with accessToken string
}

function login() {
  // login code
}

function logout() {
  // logout code
}

let provider = new SimpleProvider(getAccessToken, login, logout);
```

### Manage state

In order for the components to be aware of the state of the provider, you will need to call the `provider.setState(state: ProviderState)` method whenever the state changes. For example, when the user has logged in, call `provider.setState(ProviderState.SignedIn)`. The `ProviderState` enum defines three states:

```ts
export enum ProviderState {
  Loading,
  SignedOut,
  SignedIn
}
```

## IProvider

You can extend the `IProvider` abstract class to create your own provider.

### State

A provider must keep track of the authentication state and update the components when the state changes. The `IProvider` class already implements the `onStateChanged(eventHandler)` handler and the `state: ProviderState` property. You as a developer just need to use the `setState(state:ProviderState)` method in your implementation to update the state when it changes. Updating the state will fire the stateChanged event and update all the components automatically.

### Login/Logout

If your provider provides login or logout functionality, implement the `login(): Promise<void>` and `logout(): Promise<void>` methods. These methods are optional.

### Access Token

You must implement the `getAccessToken({'scopes': scopes[]}) : Promise<string>` method. This method is used to get a valid token before every call to the Microsoft Graph.

### Graph

The components use the Microsoft Graph Javascript SDK for all calls to the Microsoft Graph. Your provider must make the sdk available through the `graph` property. In you constructor, create a new Graph instance through

```js
this.graph = new Graph(this);
```

The `Graph` class is a light wrapper on top of the Microsoft Graph sdk.

### Example

All the providers extend the `IProvider` abstract class. Take a look at any of the [existing providers'](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/master/src/providers) source code for an example

## Set the global provider

Components use the `Providers.globalProvider` property to access a Provider. Once you have created your own provider, make sure to set this property to your provider:

```ts
import { Providers } from '@microsoft/mgt';

Providers.globalProvider = myProvider;
```

All the components will be notified of the new provider and start using it.
