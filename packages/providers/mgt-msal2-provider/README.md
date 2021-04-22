# Microsoft Graph Toolkit MSAL 2.0 Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-msal2-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-msal2-provider` package exposes the `Msal2Provider` class which uses [msal-browser](https://www.npmjs.com/package/@azure/msal-browser) to sign in users and acquire tokens to use with Microsoft Graph.


## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-msal2-provider
    ```

2. Initialize the provider in code

    ```ts
    import {Providers, LoginType} from '@microsoft/mgt-element';
    import {Msal2Provider, PromptType} from '@microsoft/mgt-msal2-provider';

    // initialize the auth provider globally
    Providers.globalProvider = new Msal2Provider({
      clientId: 'clientId',
      scopes?: string[],
      authority?: string,
      redirectUri?: string,
      loginType?: LoginType, // LoginType.Popup or LoginType.Redirect (redirect is default)
      prompt?: PromptType, // PromptType.CONSENT, PromptType.LOGIN or PromptType.SELECT_ACCOUNT
      sid?: string, // Session ID
      loginHint?: string,
      domainHint?: string,
      options?: Configuration // msal js Configuration object
    });
    ```

3. Alternatively, initialize the provider in html (only `client-id` is required):

    ```html
    <script type="module" src="../node_modules/@microsoft/mgt-msal2-provider/dist/es6/index.js" />

    <mgt-msal2-provider client-id="<YOUR_CLIENT_ID>"
                      login-type="redirect/popup" 
                      scopes="user.read,people.read" 
                      redirect-uri="https://my.redirect/uri" 
                      authority=""> 
    </mgt-msal2-provider> 
    ```

See [provider usage documentation](https://docs.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
