# Microsoft Graph Toolkit MSAL 2.0 Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-msal2-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-msal2-provider` package exposes the `Msal2Provider` class which uses [msal-browser](https://www.npmjs.com/package/@azure/msal-browser) to sign in users and acquire tokens to use with Microsoft Graph.


## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-msal2-provider
    ```

2. Initialize the provider in code with `Msal2Config`

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
      isIncrementalConsentDisabled?: boolean, //Disable incremental consent, true by default
      options?: Configuration // msal js Configuration object
    });
    ```

3. Initialize the provider in code with `Msal2PublicClientApplicationConfig` if a `PublicClientApplication` is already instantiated. For example, `msal-angular` instantiates `PublicClientApplication` on startup.

    ```ts
    import {Providers, LoginType} from '@microsoft/mgt-element';
    import {Msal2Provider, PromptType} from '@microsoft/mgt-msal2-provider';
    import {PublicClientApplication} from '@azure/msal-browser';

    // initialize the auth provider globally
    Providers.globalProvider = new Msal2Provider({
      publicClientApplication: PublicClientApplication,
      scopes?: string[],
      authority?: string,
      redirectUri?: string,
      loginType?: LoginType, // LoginType.Popup or LoginType.Redirect (redirect is default)
      prompt?: PromptType, // PromptType.CONSENT, PromptType.LOGIN or PromptType.SELECT_ACCOUNT
      sid?: string, // Session ID
      loginHint?: string,
      domainHint?: string,
      isIncrementalConsentDisabled?: boolean, //Disable incremental consent, true by default
      options?: Configuration // msal js Configuration object
    });
    ```

4. Alternatively, initialize the provider in html (only `client-id` is required):

    ```html
    <script type="module" src="../node_modules/@microsoft/mgt-msal2-provider/dist/es6/index.js" />

    <mgt-msal2-provider client-id="<YOUR_CLIENT_ID>"
                      login-type="redirect/popup" 
                      scopes="user.read,people.read" 
                      redirect-uri="https://my.redirect/uri" 
                      authority=""> 
    </mgt-msal2-provider> 
    ```
Add the `incremental-consent-disabled` boolean attribute if you wish to disable incremental consent.

See [provider usage documentation](https://learn.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)
