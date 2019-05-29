---
title: "Microsoft Teams Provider"
description: "Use the Microsoft Teams provider to enable authentication and graph access for the Microsoft Graph Toolkit components"
localization_priority: Normal
author: nmetulev
---

# Microsoft Teams provider

Use the Teams provider inside your Microsoft Teams Tab to facilitate authentication and Microsoft Graph access to all components.

Visit [the authentication docs](../providers.md) to learn more about the role of providers in the Microsoft Graph Toolkit.

## Getting started

Before using the Teams provider, you will need to make sure you have referenced the [Microsoft Teams SDK](https://docs.microsoft.com/en-us/javascript/api/overview/msteams-client?view=msteams-client-js-latest#using-the-sdk) in your page.

Here is an example using the provider in HTML (via CDN):

```html
  <!-- Microsoft Teams sdk must be referenced before the toolkit -->
  <script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" crossorigin="anonymous"></script>
  <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>

  <mgt-teams-provider
    client-id="<YOUR_CLIENT_ID>"
    auth-popup-url="https://<YOUR-DOMAIN>.com/AUTH-PATH"
  ></mgt-teams-provider>

```

Here is an example using the provider in JS modules (via NPM):

Make sure to install both the toolkit and the Microsoft Teams sdk

```bash
npm install @microsoft/mgt @microsoft/teams-js
```

Then import and use the provider

```ts
import '@microsoft/teams-js';
import {Providers, TeamsProvider} from '@microsoft/mgt'; 

Providers.globalProvider = new TeamsProvider(config);
```

where `config` is

```ts
export interface TeamsConfig {
  clientId: string;
  authPopupUrl: string;
  scopes?: string[];
  msalOptions?: Configuration;
}
```

See [sample](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/master/samples/teams-tab) for full example

## Configuring your Teams App

If you are just getting started with Teams Apps, you can follow the getting started guide [here](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/tabs/tabs-overview). You could also use the [App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-app-studio) to quickly develop your app manifest for Microsoft Teams.

Once you have installed your app with a tab, and you are ready to use the components, you will need to make sure your app has the right permissions to access the graph. Follow this 3 steps to configure your app with the necessary permissions:

1. https://docs.microsoft.com/en-us/azure/active-directory/identity-protection/graph-get-started#retrieve-your-domain-name
2. https://docs.microsoft.com/en-us/azure/active-directory/identity-protection/graph-get-started#create-a-new-app-registration
3. https://docs.microsoft.com/en-us/azure/active-directory/identity-protection/graph-get-started#grant-your-application-permission-to-use-the-api

It's important to add the right permission on the `Add API access page` as described in the links above. You will need an Administrator to add and approve the permissions, depending on which component you need.

> hint: if you are not sure what permissions to add, each component documentation includes all the permissions it needs.

### Enable Implicit Grant Flow

Make sure to enable Implicit Grant Flow as this is a requirement for web apps requesting tokens from client side (at the Azure Portal, when managing your App Registration, edit the manifest and change `oauth2AllowImplicitFlow` to `true`.

### Create the popup page

In order to login with your Teams credentials, you need to provide a URL that the Teams App with open in a popup, which will follow the authentication flow. This URL needs to be in your domain, and it needs to call the `TeamsProvider.handleAuth();` method. That's the only thing that this page needs to do. For example:

```html
<script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" crossorigin="anonymous"></script>
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js">
</script>
    mgt.TeamsProvider.handleAuth();
</script>
```

After you publish this page in your website, you need to get it's URL and use it in the `auth-popup-url/authPopupUrl` property. This URL also needs to be configured as a valid redirect URI at your app configuration, at the AAD portal.
