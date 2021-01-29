# Sample for Microsoft Graph Toolkit Electron Provider
The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-electron-provider` package exposes the `ElectronAuthenticator` and `ElectronProvider` classes which use [MSAL node](https://www.npmjs.com/package/@azure/msal-node) to sign in users and acquire tokens to use with Microsoft Graph.

This sample shows how the ElectronProvider works in a simple Electron application.

## To Use

# Navigate to this project folder
```cmd
cd /samples/src/electron-app
```
# Install dependencies
```cmd
yarn
```
# Run the app
```cmd
yarn start
```


See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.
See [Electron provider documentation](https://docs.microsoft.com/graph/toolkit/providers/electron)

## See also
* [Build an electron app and integrate Microsoft Graph Toolkit](https://docs.microsoft.com/graph/toolkit/get-started/build-an-electron-app)
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
  
