# Microsoft Graph Toolkit TeamsFx Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-teamsfx-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-teamsfx-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-teamsfx-provider` package exposes the `TeamsFxProvider` class which uses [TeamsFx](https://www.npmjs.com/package/@microsoft/teamsfx) to sign in users and acquire tokens to use with Microsoft Graph.

## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-teamsfx-provider @microsoft/teamsfx
    ```

2. Initialize the provider as below:

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {TeamsFxProvider} from '@microsoft/mgt-teamsfx-provider';
    import {TeamsUserCredential} from "@microsoft/teamsfx";

    const scope = ["User.Read"];
    const teamsfx = new TeamsFx();
    const provider = new TeamsFxProvider(teamsfx, scope);
    Providers.globalProvider = provider;
   ```

3. Perform consent process to get access token for certain scopes
    ```ts
    await teamsfx.login(this.scope);
    Providers.globalProvider.setState(ProviderState.SignedIn);
    ```


See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

You can also refer [this sample](https://github.com/OfficeDev/TeamsFx-Samples/tree/ga/graph-toolkit-contact-exporter) to learn how to use this auth provider in TeamsFx project.

## See also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
* [TeamsFx docs](https://aka.ms/teamsfx-docs)
