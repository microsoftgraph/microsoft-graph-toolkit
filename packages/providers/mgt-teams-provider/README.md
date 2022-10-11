# Microsoft Graph Toolkit Microsoft Teams Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-teams-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-teams-provider)

⚠️⚠️⚠️ This package is no longer receiving new features and will only receive critical bug and security fixes. All new applications should use [`@microsoft/mgt-teams-msal2-provider`](https://learn.microsoft.com/graph/toolkit/providers/teams-msal2) instead. ⚠️⚠️⚠️

The `@microsoft/mgt-teams-provider` package exposes the `TeamsProvider` class to be used inside your Microsoft Teams tab applications to authenticate users, to call Microsoft Graph, and to power the mgt components. The provider is built on top of [msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-core) and supports interactive authentication on the client via the Implicit Grant flow.

For Single Sing-On (SSO) support or interactive authentication based on the more secure OAuth 2.0 Authorization Code Flow with PKCE, please use the [`@microsoft/mgt-teams-msal2-provider`](https://learn.microsoft.com/graph/toolkit/providers/teams-msal2) instead.

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

[See docs for full documentation of the TeamsProvider](https://learn.microsoft.com/graph/toolkit/providers/teams)

## Usage

The TeamsProvider requires the usage of the Microsoft Teams SDK which is not automatically installed.

1. Install the packages

    ```bash
    npm install @microsoft/teams-js @microsoft/mgt-element @microsoft/mgt-teams-provider
    ```

1. Before initializing the provider, create a new page in your application (ex: https://mydomain.com/auth) that will handle the auth redirect. Call the `handleAuth` function to handle all authentication on your behalf.

    ```ts
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";
    import {TeamsProvider} from '@microsoft/mgt-teams-provider';

    TeamsProvider.microsoftTeamsLib = MicrosoftTeams;
    TeamsProvider.handleAuth();
    ```

3. Initialize the provider in your main code (not on your auth page)

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {TeamsProvider} from '@microsoft/mgt-teams-provider';
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";

    TeamsProvider.microsoftTeamsLib = MicrosoftTeams;

    Providers.globalProvider = new TeamsProvider({
      clientId: string;
      authPopupUrl: string; // ex: "https://mydomain.com/auth" or "/auth"
      scopes?: string[];
      msalOptions?: Configuration;
    })
    ```

3. Alternatively, initialize the provider in html (only `client-id` and `auth-popup-url` is required):

    ```html
    <script type="module" src="../node_modules/@microsoft/mgt-teams-provider/dist/es6/index.js" />

    <mgt-teams-provider client-id="<YOUR_CLIENT_ID>"
                        auth-popup-url="/AUTH-PATH"
                        scopes="user.read,people.read..." 
                        authority=""></mgt-teams-provider>
    ```

See [provider usage documentation](https://learn.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)