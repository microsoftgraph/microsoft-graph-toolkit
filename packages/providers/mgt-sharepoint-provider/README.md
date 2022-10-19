# Microsoft Graph Toolkit SharePoint Provider

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-sharepoint-provider?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-sharepoint-provider)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph. 

The `@microsoft/mgt-sharepoint-provider` package exposes the `SharePointProvider` class to be used inside your SharePoint web parts to power the components with Microsoft Graph access.

[See docs for full documentation of the SharePointProvider](https://learn.microsoft.com/graph/toolkit/providers/sharepoint)

## Usage

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-sharepoint-provider
    ```

2. Initialize the provider in code

    ```ts
    import {Providers} from '@microsoft/mgt-element';
    import {SharePointProvider} from '@microsoft/mgt-sharepoint-provider';

    // add the onInit() method if not already there in your web part class
    protected async onInit() {
        Providers.globalProvider = new SharePointProvider(this.context);
    }
    ```

See [provider usage documentation](https://learn.microsoft.com/graph/toolkit/providers) to learn about how to use the providers with the mgt components, to sign in/sign out, get access tokens, call Microsoft Graph, and more.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)