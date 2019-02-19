# Authentication and Graph helpers

## Description
The M365 Community Toolkit contains authentication and graph helpers and interfaces. This enables all controls to be aware of the current authentication state and initiate calls to the graph without the developer having to write the code themselves. 

## Authentication
The authentication helpers can be used to provide global authentication context to all the controls. This allows the controls to be aware of the current state without explicitly having to write any code for each control.

If you do not already have an application registered with the Microsoft identity platform (also known as Azure Active Directory), [make sure to that first](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app). Make sure to enable Implicit Grant Flow as this is a requirement for web apps requesting tokens from client side. Once you have a `ClientID`, and you've selected the scopes/permissions in the Azure portal, you should be good to go.

### Examples

These are examples of initializing an authentication provider in JavaScript. You can also initialize an authentication provider in HTML through the [my-login](./controls/login-control.md) control

1. Initialize for MSAL

    in JS:
    ```js
    Auth.initMSALProvider("client-id");
    ```

2. Initialize with existing MSAL instance

    only in JS:
    ```js
    Auth.initWithMSAL(MyUserAgentApplication)
    ```

3. Initialize for Progressive Web Apps on Windows

    \\ TODO

4. Initialize for Teams

    \\ TODO

5. Initialize for SharePoint

    \\ TODO

6. Initialize for AddIns

    \\ TODO

### API

1. The `Auth` global variable exposes the following functions
    * `function initMSALProvider(config : MSALConfig)`
        
        where `MSALConfig` defines the following interface:

        ```ts
        interface MSALConfig {

            clientId: string; // required
            scopes?: string[]; // optional - default = ['user.read']
            authority?: string; // optional - default = null
            loginType?: LoginType; // optional - default = Redirect
            options?: any; // optional - default = null
        }
        ```

        Visit the MSAL.js documentation for [authority](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics#initialization-of-msal) and [options](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki/MSAL-basics#configuration-options)

        In general, you will not need to specify the scope. Each control will automatically request access to the appropriate scopes needed for the control to function properly. Make sure to visit each control's documentation to know what scope to enable for your app in the Azure Portal

    * `function initWithMSAL(application : UserAgentApplication)`

        If you are already using MSAL.js in your application, just pass in the instance to the `initWithMSAL` method. No need to change how you configure your application

    * `function initForWinPWA(clientId: string)`

        If your application will be running as a Progressive Web App (PWA) on Windows, you might want to take advantage of the native [Web Account Manager (WAM)](https://docs.microsoft.com/en-us/windows/uwp/security/web-account-manager) APIs. This will allow you to simplify the login process for your users by using the built in Windows identity provider.

    * `function getAuthProvider() : IAuthProvider`

        This function is intended to be used by each control to get a reference to the authentication provider once initialized. The provider defines the following interface:

        ```ts
        export interface IAuthProvider 
        {
            type : ProviderType;
            readonly isLoggedIn : boolean;

            login() : Promise<void>;
            logout() : Promise<void>;
            getAccessToken(scopes?: string[]) : Promise<string>;

            // get access to underlying provider (such as UserAgentApplication)
            provider : any;

            // graph API - see documentation bellow
            graph : Graph;

            // calls the callback function when login state changes
            onLoginChanged(eventHandler : EventHandler<LoginChangedEvent>)
        }
        ```

    * `function onAuthProviderChanged(callbackFunction)`

        the `callbackFunction` function will be called when a new provider is initialized. the new provider will be passed to the function.

## Microsoft Graph

Each control is enabled with Graph access out of the box as long as the developer has initialized a provider (as described in the above section). To get a reference to the built in Microsoft Graph APIs, first get a reference to the global IAuthProvider and then use the `Graph` object:

```js
let provider = Auth.getAuthProvider();
if (provider) {
    let graph = provider.Graph;
}
```

The Graph object defines several functions for accessing the Graph:

* `get(resource: string, scopes?: string[]) : Promise<Response>`

* `getJson(resource: string, scopes? : string[]) :Promise<any>`

Use these two functions to get data from any graph resource (such as `'/me'`).

If you do not pass in the scope (such as `['user.read']` for `'/me'`), the call will use the scopes you specified when creating the IAuthProvider (or default values).

> NOTE: If you know all scopes your application will use, it might be easier to just initialize the `IAuthProvider` with those scopes once and not worry about what scopes to use for each call. See documentation above.

The Graph object also provides common functions for resources needed by  controls. For example, `graph.me()` will get the `'/me'` resource. For list of all functions, see the API documentation (TODO);

