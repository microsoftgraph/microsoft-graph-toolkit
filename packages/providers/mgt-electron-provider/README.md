# Microsoft Graph Toolkit Electron Provider
The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-electron-provider` package exposes the `ElectronAuthenticator` and `ElectronProvider` classes which use [MSAL node](https://www.npmjs.com/package/@azure/msal-node) to sign in users and acquire tokens to use with Microsoft Graph.


## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-electron-provider
    ```

2. Initialize the provider in your renderer process (Front end, eg. renderer.ts)

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {ElectronProvider} from '@microsoft/mgt-electron-provider/dist/ElectronProvider';

    // initialize the auth provider globally
    Providers.globalProvider = new ElectronProvider();
    ```

3. Initialize ElectronAuthenticator in Main.ts (Back end)

    ```ts
    import { ElectronAuthenticator } from '@microsoft/mgt-electron-provider/dist/ElectronAuthenticator';

    const authProvider = new ElectronAuthenticator({
      clientId: '[client-id]]',
      authority: '[authority-url]',
      mainWindow: mainWindow //Main window on which you would want to authenticate the user
      scopes: ['User.Read'], //optional
    });
    ```

See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.
See [Electron provider documentation](https://docs.microsoft.com/graph/toolkit/providers/electron)

## See also
* [Build an electron app and integrate Microsoft Graph Toolkit](https://docs.microsoft.com/graph/toolkit/get-started/build-an-electron-app)
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
  
