# Providers

The Microsoft Graph Toolkit defines Providers that enable all components to be aware of the current authentication state and initiate calls to the Microsoft Graph without the developer having to write the code themselves. Each provider provides implementation for acquiring the necessary access token for calling the Microsoft Graph APIs.

In order for the components to use a provider, the `Providers.globalProvider` property must be set to a Provider you'd like to use. 

Example of using the MsalProvider:
```js
Providers.globalProvider = new MsalProvider({
    clientId: '[CLIENT_ID]'
});
```

The toolkit implements several providers:
* [MsalProvider](./providers/msal.md)
* [SharePointProvider](./providers/sharepoint.md)
* [TeamsProvider](./providers/teams.md)
* Office Add-ins provider (coming soon)


## Get started

You can create a provider at any time, but it's recommended to create it before using any of the components. 

### API

1. The `Providers` global variable exposes the following properties and functions
    * `globalProvider : IProvider`

        set this property to a provider you want to use globally. All components use this property to get a reference to the provider. Setting this property will fire the onProvidersChanged event

    * `function onProviderUpdated(callbackFunction)`

        the `callbackFunction` function will be called when a provider is changed or when the state of a Provider changes. A `ProvidersChangedState` enum value will be passed to the function to indicate what updated

## Implement your own provider

The toolkit provides two ways to create new providers:

* Create a new `SimpleProvider` by passing in a function for getting an access token, or
* Extend the `IProvider` abstract class

Read more about each one in the [custom providers](./providers/custom.md) documentation;

## Making your own calls to the Microsoft Graph

All components can access the Microsoft Graph out of the box as long as the developer has initialized a provider (as described in the above section). To get a reference to the same Microsoft Graph SDK used by the components, first get a reference to the global IProvider and then use the `Graph` object:

```js
let provider = Providers.globalProvider;
if (provider) {
    let graphClient = provider.graph.client;
    let userDetails = await graphClient.api("me").get();
}
```

There might be cases were you will need to pass additional scopes depending on the api you are calling:

```js
graphClient.api('me')
           .middlewareOptions([{
                scopes: ['user.read']
            }]).get();
```



The graph object is an instance of the [Microsoft Graph Javascript SDK](https://github.com/microsoftgraph/msgraph-sdk-javascript) and you can use it to make any calls to the Microsoft Graph.

