# Authentication and Graph helpers

## Description
The Graph Toolkit contains authentication and graph helpers and interfaces. This enables all components to be aware of the current authentication state and initiate calls to the graph without the developer having to write the code themselves. 

## Authentication
The authentication helpers can be used to provide global authentication context to all the controls. This allows the controls to be aware of the current state without explicitly having to write any code for each control.

If you do not already have an application registered with the Microsoft identity platform (also known as Azure Active Directory), [make sure to that first](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app). Make sure to enable Implicit Grant Flow as this is a requirement for web apps requesting tokens from client side. Once you have a `ClientID`, and you've selected the scopes/permissions in the Azure portal, you should be good to go.

### Examples

These are examples of initializing an authentication provider in JavaScript. 

1. Initialize for MSAL

    ```js
    Providers.globalProvider = new MsalProvider({clientId: '<YOUR_CLIENT_ID>'});
    ```

2. Initialize for Progressive Web Apps on Windows

    ```js
    Providers.globalProvider = new WamProvider('<YOUR_CLIENT_ID>');
    ```

3. Initialize for Teams

    ```js
    Providers.globalProvider = new TeamsProvider('<YOUR_CLIENT_ID>', '<LOGIN_REDIRECT_URL>');
    ```
    For more see [Teams Provider docs](providers/Teams.md).

4. Initialize for SharePoint

    ```js
    Providers.globalProvider = new SharePointProvider(this.context); 
    ```
    For more see [Sharepoint Provider docs](providers/SharePoint.md).

5. Initialize for AddIns

    \\\\ TODO

### Examples (HTML)

Most providers also have a web component wrapper that allows you to initialize it all in HTML

1. Use the `mgt-msal-provider` for MSAL. See [docs](./components/providers/mgt-msal-provider.md)

    ```html
    <mgt-msal-provider client-id="<YOUR_CLIENT_ID>"></mgt-msal-provider>
    ```

2. Use the `mgt-wam-provider` for authenticating with the Web Authentication Manager in Windows. See [docs](./components/providers/mgt-wam-provider.md)

    ```html
    <mgt-wam-provider client-id="<YOUR_CLIENT_ID>"></mgt-wam-provider>
    ```

## Get started

You can initialize a provider at any time, but it's recommended to initialize them before using any of the components. 

### API

1. The `Providers` global variable exposes the following properties and functions
    * `GlobalProvider : IProvider`

        set this property to a provider you want to use globally. All components use this property to get a reference to the provider. Setting this property will fire the onProvidersChanged event

    * `function onProvidersChanged(callbackFunction)`

        the `callbackFunction` function will be called when a new provider is set. Check the `GlobalProvider` property to get the new provider.

## Implement your own provider

You will need to extend the `IProvider` abstract class to create your own provider.

### State
A provider must keep track of the authentication state and update the components when the state changes. The `IProvider` implements the `onStateChanged(eventHandler)` handler and the `state:ProviderState` property. Use the `setState(state:ProviderState)` method from your implementation to update the state. Updating the state will fire the stateChanged event and update all the components automatically.

### Login/Logout
If your provider provides login or logout functionality, implement the `login() : Promise<void>` and `logout() : Promise<void>` methods. These methods are optional.

### Access Token
You must implement the `getAccessToken(...scopes: string[]) : Promise<string>` method. This method is used to get a valid token before every call to the Microsoft Graph.

### Graph
The components use the Graph sdk from each provider. Make sure to set the `Graph` property in your initialization code by calling `this.graph = new Graph(this)`.

## Microsoft Graph

Each control is enabled with Graph access out of the box as long as the developer has initialized a provider (as described in the above section). To get a reference to the built in Microsoft Graph APIs, first get a reference to the global IProvider and then use the `Graph` object:

```js
let provider = Providers.globalProvider;
if (provider) {
    let graph = provider.Graph;

    // TODO - update this when switched to GraphSDK
}
```

The graph object is an instance of the [GraphSDK](https://github.com/microsoftgraph/msgraph-sdk-javascript) and you can use it to make any calls to the Microsoft Graph. It is also used by all the components.

