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
    GraphProviders.addMsalProvider({clientId: '<YOUR_CLIENT_ID>'});
    ```

2. Initialize for Progressive Web Apps on Windows

    ```js
    GraphProviders.addWamProvider('<YOUR_CLIENT_ID>');
    ```

3. Initialize for Teams

    ```js
    await Providers.addTeamsProvider('<YOUR_CLIENT_ID>', '<LOGIN_REDIRECT_URL>');
    ```
    For more see [Teams Provider docs](providers/Teams.md).

4. Initialize for SharePoint

    ```js
    Providers.addSharePointProvider(this.context); 
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

2. Use the 'mgt-wam-provider' for authenticating with the Web Authentication Manager in Windows. See [docs](./components/providers/mgt-wam-provider.md)

    ```html
    <mgt-wam-provider client-id="<YOUR_CLIENT_ID>"></mgt-wam-provider>
    ```

## Get started

You can initialize a provider at any time, but it's recomended to initialize them before using any of the components. You can also initialize multiple providers at the same time when you need to target multiple platforms. For example, if you initialize both the Wam and Msal providers (in that order), the components will chose the first provider available on the current runtime. When running in a progressive web app (PWA) on Windows, the Wam provider will be used since it is available and was initialized first. When running in the browser, Wam is not available, so the Msal provider will be used.

### API

1. The `Providers` global variable exposes the following functions
    * `function addMsalProvider(config : MsalConfig)`
        
        where `MsalConfig` defines the following interface:

        ```ts
        interface MsalConfig {

            clientId: string; // required
            scopes?: string[]; // optional - default = ['user.read']
            authority?: string; // optional - default = null
            loginType?: LoginType; // optional - default = Redirect
            options?: any; // optional - default = null
        }
        ```

        Visit the MSAL.js documentation for [authority](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics#initialization-of-msal) and [options](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics#configuration-options)

        In general, you will not need to specify the scope. Each control will automatically request access to the appropriate scopes needed for the control to function properly. Make sure to visit each control's documentation to know what scope to enable for your app in the Azure Portal

    * `function addWamProvider(clientId : string, authority? : string)`

        If your application will be running as a Progressive Web App (PWA) on Windows, you might want to take advantage of the native [Web Account Manager (WAM)](https://docs.microsoft.com/en-us/windows/uwp/security/web-account-manager) APIs. This will allow you to simplify the login process for your users by using the built in Windows identity provider.

    * `function addSharePointProvider(context : WebPartContext)`

        See [docs](providers/SharePoint.md)

    * `async function addTeamsProvider(clientId : string, loginPopupUrl : string)`

        TODO

    * `function addCustomProvider(provider : IProvider)`

        Use the add function to initialize with a custom provider

    * `function getAvailable() : IProvider`

        This function is intended to be used by each control to get a reference to the authentication provider once initialized. It returns the first provider that is available on the current runtime.

    * `function onProvidersChanged(callbackFunction)`

        the `callbackFunction` function will be called when a new provider is initialized. Use `getAvailable()` to get the first available provider.

## Implement your own provider

The provider defines the following interface:

```ts
export interface IProvider 
{
    readonly isLoggedIn : boolean;

    login?() : Promise<void>;
    logout?() : Promise<void>;
    getAccessToken(...scopes: string[]) : Promise<string>;

    // get access to underlying provider (such as UserAgentApplication)
    provider : any;

    // graph API - see documentation bellow
    graph : Graph;

    // calls the callback function when login state changes
    onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>)
}
```

## Microsoft Graph

Each control is enabled with Graph access out of the box as long as the developer has initialized a provider (as described in the above section). To get a reference to the built in Microsoft Graph APIs, first get a reference to the global IProvider and then use the `Graph` object:

```js
let provider = GraphProviders.getAvailable();
if (provider) {
    let graph = provider.Graph;
}
```

The graph object is an instance of the [GraphSDK](https://github.com/microsoftgraph/msgraph-sdk-javascript) and you can use it to make any calls to the Microsoft Graph. It is also used by all the components.

