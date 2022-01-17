# Microsoft Graph Toolkit TeamsFx Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-teamsfx-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-teamsfx-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-teamsfx-provider` package exposes the `TeamsFxProvider` class which uses [TeamsFx](https://www.npmjs.com/package/@microsoft/teamsfx) to sign in users and acquire tokens to use with Microsoft Graph.

## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-teamsfx-provider @microsoft/teamsfx
    ```

2. Initialize the provider in code with `TeamsFxConfig`

    ```ts
    import {Providers, LoginType} from '@microsoft/mgt-element';
    import {TeamsFxProvider} from '@microsoft/mgt-teamsfx-provider';
    import {TeamsUserCredential} from "@microsoft/teamsfx";

    const scope = ["User.Read"];
    const credential = new TeamsUserCredential();
    const provider = new TeamsFxProvider(credential, scope);
    Providers.globalProvider = provider;
   ```

See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## See also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)