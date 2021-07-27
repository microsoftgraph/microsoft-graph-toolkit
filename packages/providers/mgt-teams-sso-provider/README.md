# Microsoft Graph Toolkit Microsoft Teams SSO Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-teams-sso-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-teams-sso-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-teams-sso-provider` package exposes the `TeamsSSOProvider` class to be used inside your Microsoft Teams tab applications to authenticate users, to call Microsoft Graph, and to power the mgt components.

[See docs for full documentation of the TeamsSSOProvider](https://docs.microsoft.com/graph/toolkit/providers/teamssso)

## Usage

The TeamsProvider requires the usage of the Microsoft Teams SDK which is not automatically installed.

1. Install the packages

    ```bash
    npm install @microsoft/teams-js @microsoft/mgt-element @microsoft/mgt-teams-sso-provider
    ```

1. Before initializing the provider, create a new page in your application (ex: https://mydomain.com/auth) that will handle the auth redirect. Call the `handleAuth` function to handle all client side auth or permission consent.

    ```ts
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";
    import {TeamsSSOProvider} from '@microsoft/mgt-teams-sso-provider';

    TeamsProvider.microsoftTeamsLib = MicrosoftTeams;
    TeamsProvider.handleAuth();
    ```

3. Initialize the provider in your main code (not on your auth page). The provider can be used in "client side auth" mode or SSO mode. SSO mode is enabled by setting `ssoUrl` \ `sso-url` and requires a backend service to handle the on-behalf-of flow.

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {TeamsSSOProvider} from '@microsoft/mgt-teams-sso-provider';
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";

    TeamsProvider.microsoftTeamsLib = MicrosoftTeams;

    Providers.globalProvider = new TeamsSSOProvider({
      clientId: string;
      authPopupUrl: string; // ex: "https://mydomain.com/auth" or "/auth"
      scopes?: string[];
      msalOptions?: Configuration;
      ssoUrl?: string; // ex: '/api/token',
      autoConsent?: boolean,
      httpMethod: HttpMethod; //ex HttpMethod.POST
    })
    ```

3. Alternatively, initialize the provider in html (only `client-id` and `auth-popup-url` is required):

    ```html
    <script type="module" src="../node_modules/@microsoft/mgt-teams-provider/dist/es6/index.js" />

    <mgt-teams-sso-provider client-id="<YOUR_CLIENT_ID>"
                        auth-popup-url="/AUTH-PATH"
                        scopes="user.read,people.read..." 
                        authority=""
                        sso-url="/api/token" 
                        http-method="POST">
                        ></mgt-teams-provider>
    ```

See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)