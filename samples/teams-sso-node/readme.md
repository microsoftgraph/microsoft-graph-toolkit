# Node.js Teams Single Sign On (SSO) Sample
This sample is using the Microsoft Graph Toolkit's [Teams SSO Provider](https://docs.microsoft.com/graph/toolkit/providers/teamssso) and it will:
1. Get an access token for the logged-in user via the Microsoft Teams SSO experience, with the help of the Teams SDK.
2. Call the Node.js backend to exchange the current token for one with the requested scopes - using `on-behalf-of flow`
3. Pass the access token to the MGT components in order to make Microsoft Graph API calls. 

## On-behalf-of

In on-behalf-of mode, the service uses the Microsoft identity platform's [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the token sent by the calling client for a Microsoft Graph token. The next section gives instructions on how to register the app in Azure Active Directory.

## Prerequisites

You will need:

1. A global administrator account for an Office 365 tenant. Testing in a production tenant is not recommended! You can get a free tenant for development use by signing up for the [Office 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).

1. To test locally, [NodeJS](https://nodejs.org/en/download/) must be installed on your development machine.

1. To test locally, you'll need [Ngrok](https://ngrok.com/) installed on your development machine.
Make sure you've downloaded and installed Ngrok on your local machine. ngrok will tunnel requests from the Internet to your local computer and terminate the SSL connection from Teams.

> NOTE: The free ngrok plan will generate a new URL every time you run it, which requires you to update your Azure AD registration, the Teams app manifest, and the project configuration. A paid account with a permanent ngrok URL is recommended.


## App registration
Your tab needs to run as a registered Azure AD application to obtain an access token from Azure AD. In this step, you'll register the app in your tenant and give Microsoft Teams permission to obtain access tokens on its behalf.

1. Run Ngrok to expose your local web server via a public URL. Make sure to point it to your Ngrok URI. This sample uses port `5000` locally. Run: 
    * No custom sub domain: `./ngrok http 5000`
    * Custom sub domain: `/ngrok http 5000 -subdomain="mgtsso"`

Leave this running while you're running the application locally, and open another command prompt for the steps which follow.

1. Create an [AAD application](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#1-create-your-aad-application-in-azure) in Azure. You can do this by visiting the "Azure AD app registration" portal in Azure. Make sure to copy the `Application Id` and the `Tenant Id`, you will use them later.

    * Set your application URI to the same URI you've created in Ngrok. 
        * Ex: `api://mgtsso.ngrok.io/{appId}`
        using the application ID that was assigned to your app
    * Setup your redirect URIs. This will allow Azure AD to return authentication results to the correct URI. This is required for consent to work.
        * Visit `Manage > Authentication`. 
        * Create a redirect URI in the format of: `https://mgtsso.ngrok.io`.
        * Enable Implicit Grant by selecting `Access Tokens` and `ID Tokens`.
    * Setup a client secret. You will need this when you exchange the token for more API permissions from your backend.
        * Visit `Manage > Certificates & secrets`
        * Create a new client secret.
    * Setup your API permissions. This is what your application is allowed to request permission to access.
        * Visit `Manage > API Permissions`
        * Make sure you have the following Graph permissions enabled: `email`, `offline_access`, `openid`, `profile`, and `User.Read`.
    * Expose an API that will give the Teams desktop, web and mobile clients access to the permissions above
        * Visit `Manage > Expose an API`
        * Add a scope and give it a scope name of `access_as_user`. Your API url should look like this: `api://mgtsso.ngrok.io/{appID}/access_as_user`. In the "who can consent" step, enable it for "Admins and users". Make sure the state is set to "enabled".
        * Next, add two client applications. This is for the Teams desktop/mobile clients and the web client.
            * 5e3ce6c0-2b1f-4285-8d4b-75ee78787346
            * 1fec8e78-bce4-4aaf-ab1b-5451cc387264

## Create Teams App
1. Now you can use [App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-app-studio) to quickly develop your app manifest for Microsoft Teams and test the app

* You could look at the `example.manifest.json` as inspiration. You would have to change the following in this manifest file:
    * Generate a new unique ID for the application and replace the id field with this GUID. On Windows, you can generate a new GUID in PowerShell with this command:
    ~~~ powershell
     [guid]::NewGuid()
    ~~~
    * Ensure the package name is unique within the tenant where you will run the app
    * Replace `{ngrokSubdomain}` with the subdomain you've assigned to your Ngrok account in step #1 above.
    * Update your `webApplicationInfo` section with your Azure AD application ID that you were assigned in step #2 above.

## Configuring the sample

1. Rename the example.env file to .env, and set the values as follows.

    | Setting | Value |
    |---------|-------|
    | SSO_API | Set to the API URL you created in the app registration `api://mgtsso.ngrok.io/{appID}` |
    | APP_TENANT_ID | The tenant ID from your app registration |
    | APP_SECRET | The client secret from your app registration |

## Run the sample
The `Teams SSO Provider` needs the `client id` of your app registration, the `scopes` your app will use, and a relative or absolute `sso url`.
The `index.html` page displays two ways of using the provider:
- Commented away, is the way to use the component provider.
- Inside of the script tag is the way to do this in code.

### Update the sample
1. Update the `client id` to your app registration id in `index.html` 

1. Open your command-line interface (CLI) in the root of this project and run the following command to install dependencies.

    ```Shell
    yarn install
    ```

    > **NOTE**
    > This only needs to be done once.

1. Run the following command to start the sample.

      ```Shell
      yarn start
      ```

1. Open the installed Teams tab in Microsoft Teams. Using the browser enables you to debug. 

### How do I know it worked?

If everything worked correctly, the components should load data just as they would normally after a login.
If you haven't consented the scopes, you might be presented with a consent popup where you need to enter your password and then consent. After you have consented, the components should be filled with content.