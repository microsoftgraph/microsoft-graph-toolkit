# Microsoft Graph Toolkit Electron Provider
The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-electron-provider` package exposes the `ElectronAuthenticator` and `ElectronProvider` classes which uses MSAL node to sign in users and acquire tokens to use with Microsoft Graph.


## Usage

1. Install the packages

# <to be added>

2. Initialize the provider in renderer.ts (Front end)

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {ElectronProvider} from '@microsoft/mgt-electron-provider/dist/es6/ElectronProvider';

    // initialize the auth provider globally
    Providers.globalProvider = new ElectronProvider();
    ```

3. Initialize ElectronAuthenticator in Main.ts (Back end)

    ```ts
    import { ElectronAuthenticator } from '@microsoft/mgt-electron-provider/dist/es6/ElectronAuthenticator';

    const authProvider = new ElectronAuthenticator({
      clientId: '[client-id]]',
      authority: '[authority-url]',
      mainWindow: mainWindow //Main window on which you would want to authenticate the user
      scopes: ['User.Read'], //optional
    });
    ```

See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## See also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
