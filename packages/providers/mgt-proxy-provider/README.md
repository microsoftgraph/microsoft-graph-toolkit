# Microsoft Graph Toolkit Proxy Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-proxy-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-proxy-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-proxy-provider` package exposes the `ProxyProvider` class which allows a developer to proxy all calls to Microsoft Graph to their own backend. This allows all mgt-components to function properly when the authentication and calls to Microsoft Graph must be done in the backend.

[See docs for full documentation of the ProxyProvider](https://learn.microsoft.com/graph/toolkit/providers/proxy)

## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-proxy-provider
    ```

2. Initialize the provider in code

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {ProxyProvider} from '@microsoft/mgt-proxy-provider';

    // initialize the auth provider globally
    Providers.globalProvider = new ProxyProvider("https://myurl.com/api/GraphProxy");
    ```

3. Alternatively, initialize the provider in html (only `client-id` is required):

    ```html
    <script type="module" src="../node_modules/@microsoft/mgt-proxy-provider/dist/es6/index.js" />

    <mgt-proxy-provider graph-proxy-url="https://myurl.com/api/GraphProxy"></mgt-proxy-provider>
    ```

See [provider usage documentation](https://learn.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)