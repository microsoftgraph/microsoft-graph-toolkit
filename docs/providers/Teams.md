# Teams provider

Use the Teams provider inside your Microsoft Teams App, inside a Tab, to power it with Microsoft Graph access.

Visit [the authentication docs](../providers.md) to learn more about the role of providers in the Microsoft Graph Toolkit

## Getting started

Initialize the provider inside your html using the `mgt-teams-provider` component.

```html
<mgt-teams-provider
  client-id="<YOUR_CLIENT_ID>"
  auth-popup-url="https://<YOUR-DOMAIN>.com/AUTH-PATH"
></mgt-teams-provider>
```

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

In order to login with your Teams credentials, you need to provide a URL that the Teams App with open in a popup, which will follow the authentication flow. This URL needs to be in your domain, and it needs to call the `TeamsProvider.handleAuth();` method. That's the only thing that this page needs to do. For example (if you are using parcel to bundle your app) - `auth.html`:

```html
<html>
  <head>
    <script src="./node_modules/microsoft-graph-toolkit/dist/es6/index.js"></script>
    <script>
      var teamsProvider = parcelRequire('node_modules/microsoft-graph-toolkit/dist/es6/index.js');
      teamsProvider.TeamsProvider.handleAuth();
    </script>
  </head>
</html>
```

### Configure Redirect URIs

After you publish this page in your website, you need to get it's URL and use it in the `auth-popup-url` property. This URL also needs to be configured as a valid redirect URI at your app configuration, at the AAD portal.
