# Msal2Provider SSO with ASP.NET Core Web App

This sample demonstrates how to authenticate users in your ASP.NET Core Web App while also leveraging MGT components on the client. The Msal2Provider supports SSO, allowing a user to authenticate silently if they are already authenticated once, avoiding the need to authenticate twice.

## Run the sample

To run the sample, you will need to create a client id and secret first and update the code.

1. Navigate to https://aad.portal.azure.com

1. Click on **Azure Active Directory** -> **App registrations**

1. Click on **New registration**

1. Give the app a name

1. Select **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)** for the supported account type

1. Select **Web** for the redirect URI and enter `https://localhost/signin-oidc`. Click **Register**

1. Under **Manage**, select **Authentication**

1. Click on **Add a platform** and select **Single-page application**

1. Use `https://localhost/sso.html` for the redirect uri. Click **Configure**

1. Under **Manage**, select **Certificates & secrets** > **New client secret**.

1. Enter a Description and expiration and click **Add**.

1. Copy the secret and replace the placeholder in **appsettings.json** with this value.

1. Under **Overview**, find the **Application (client) ID**. Replace the client id in **appsettings.json** with your client id

1. Now you can run the app. The browser will open and ask you to sign in a user. Once you've authenticated user and consented to all scopes, you should be signed in to the web app and see several MGT components functional on the home page.

## How it works

The web app is built with ASP.NET Core and users are authenticated with the [Microsoft.Identity.Web](https://www.nuget.org/packages/Microsoft.Identity.Web) package. This allows developers to authenticate users and authorize access to Web APIs (like Microsoft Graph) from the backend. 

To use MGT components in the app, the client must also be able auth users and fetch tokens for Microsoft Graph. To avoid signing in the user twice just to be able to call Microsoft Graph from the client, we can leverage SSO to automatically sign in users via the current session that was established when the user signed in once.

First, we reference the toolkit and instantiate a new Msal2Provider in the **Pages/Shared/_Layout.cshtml** page:

```js
@if (User.Identity.IsAuthenticated)
{
    <script src="https://unpkg.com/@@microsoft/mgt/dist/bundle/mgt-loader.js"></script>

    <script>
        mgt.Providers.globalProvider = new mgt.Msal2Provider({
            clientId: "@Configuration["AzureAd:ClientId"]",
            loginHint: "@User.Claims.FirstOrDefault(c => c.Type == "preferred_username")?.Value",
            redirectUri: "/sso.html",
            loginType: mgt.LoginType.Popup
        });
    </script>
}
```

To enable SSO, we need to provide a `loginHint` (the email for the current signed in user) or an `sid` (session id). This will set the provider in SSO mode which will attempt to sign in the user if they are already signed in elsewhere. 

Notice, we are also setting a redirect Uri to an empty page we created (`/sso.html`) as page for the sso process to redirect to. We've also added this page (`https://localhost/sso.html`) as a SPA redirect uri for our AAD app.

Also notice that we are not setting any scopes here. All scopes are added on the .Net side in **appsettings.json** so that the user consents once when they sign in for the first time. 

Finally, we are now able to use the mgt components through out our app. Here are the components being used in **Index.cshtml**:

```js
@if (User.Identity.IsAuthenticated)
{
    <div>Signed in User: <mgt-person person-query="me" view="oneline"></mgt-person></div>
    
    <mgt-people-picker></mgt-people-picker>

    <mgt-agenda></mgt-agenda>
}
```
