# Microsoft Graph Toolkit Microsoft Teams Msal2 Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-teams-msal2-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-teams-msal2-provider)

The `@microsoft/mgt-teams-msal2-provider` package exposes the `TeamsMsal2Provider` class to be used inside your Microsoft Teams tab applications to authenticate users, to call Microsoft Graph, and to power the Microsoft Graph Toolkit components. The provider is built on top of [msal-browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser) and supports both the interactive sign in flow on the client and Single Sign-On (SSO) flow via your own backend. SSO mode is enabled by setting `ssoUrl` \ `sso-url` and requires a backend service to handle the on-behalf-of flow.

[See the full documentation of the TeamsMsal2Provider](https://learn.microsoft.com/graph/toolkit/providers/teams-msal2)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

## Usage

The TeamsMsal2Provider requires the usage of the Microsoft Teams SDK which is not automatically installed.

1. Install the packages

    ```bash
    npm install @microsoft/teams-js @microsoft/mgt-element @microsoft/mgt-teams-msal2-provider
    ```

1. Before initializing the provider, create a new page in your application (ex: https://mydomain.com/auth) that will handle the auth redirect. Call the `handleAuth` function to handle all client side auth or permission consent.

    ```ts
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";
    import {TeamsMsal2Provider} from '@microsoft/mgt-teams-msal2-provider';

    TeamsMsal2Provider.microsoftTeamsLib = MicrosoftTeams;
    TeamsMsal2Provider.handleAuth();
    ```

3. Initialize the provider in your main code (not on your auth page). The provider can be used in "client side auth" mode or SSO mode. SSO mode is enabled by setting `ssoUrl` \ `sso-url` and requires a backend service to handle the on-behalf-of flow.

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {TeamsMsal2Provider} from '@microsoft/mgt-teams-msal2-provider';
    import * as MicrosoftTeams from "@microsoft/teams-js/dist/MicrosoftTeams";

    TeamsMsal2Provider.microsoftTeamsLib = MicrosoftTeams;

    Providers.globalProvider = new TeamsMsal2Provider({
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

    <mgt-teams-msal2-provider client-id="<YOUR_CLIENT_ID>"
                        auth-popup-url="/AUTH-PATH"
                        scopes="user.read,people.read..." 
                        authority=""
                        sso-url="/api/token" 
                        http-method="POST">
                        ></mgt-teams-provider>
    ```

See [provider usage documentation](https://learn.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
