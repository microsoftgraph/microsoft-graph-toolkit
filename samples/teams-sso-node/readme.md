# Node.js Teams Single Sign On (SSO) Sample
This sample is using the Microsoft Graph Toolkit's [Teams Provider](https://docs.microsoft.com/graph/toolkit/providers/teams) and it will:
1. Get an access token for the logged-in user via the Microsoft Teams SSO experience, with the help of the Teams SDK.
2. Call the Node.js backend to exchange the current token for one with the requested scopes - using `on-behalf-of flow`
3. Pass the access token to the MGT components to make Microsoft Graph API calls. 

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

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Node.js Teams SSO`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to your domain. ex `https://mgtsso.ngrok.io`.
        - Enable Implicit Grant by selecting `Access Tokens` and `ID Tokens`

    > **IMPORTANT** 
    Setting up the redirect URI will allow Azure AD to return authentication results to the correct URI. This is required for consent to work.

1. Select **Register**. On the **Node.js Teams SSO** page, copy the value of the **Application (client) ID** and **Directory (tenant) ID**. Save them, you will need them in the next step.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the next step.

    > **IMPORTANT**
    > This client secret is never shown again, so make sure you copy it now.

1. Select **API permissions** under **Manage**, then select **Add a permission**.

1. Select **Microsoft Graph**, then **Delegated permissions**.

1. Make sure you have the following Graph permissions enabled: `email`, `offline_access`, `openid`, `profile`, and `User.Read`.

1. Select **Expose an API**. Select the **Set** link next to **Application ID URI**. You will be presented with `api://{Your App Id}`. Change this by adding the URI you created in Ngrok. 
    - Ex: `api://mgtsso.ngrok.io/{Your App Id}`

1. In the **Scopes defined by this API** section, select **Add a scope**. Fill in the fields as follows and select **Add scope**.

    - **Scope name:** `access_as_user`
    - **Who can consent?: Admins and users**
    - **Admin consent display name:** `Teams can access the user’s profile`
    - **Admin consent description:** `Allows Teams to call the app’s web APIs as the current user`
    - **User consent display name:** `Teams can access the user profile and make requests on the user's behalf`
    - **User consent description:** `Enable Teams to call this app’s APIs with the same rights as the user.`
    - **State: Enabled**
    
    Your API URL should look like this: `api://mgtsso.ngrok.io/{appID}/access_as_user`. 

1. Next, add two client applications. This is for the Teams desktop/mobile clients and the web client. In the **Authorized client applications** section, select **Add a client application**. Fill in the Client ID and select the scope we created. Then select **Add application**. Do this for the followings Ids
    
    - 5e3ce6c0-2b1f-4285-8d4b-75ee78787346
    - 1fec8e78-bce4-4aaf-ab1b-5451cc387264

    
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
The `Teams Provider` needs the `client id` of your app registration, the `scopes` your app will use, and a relative or absolute `sso url`.
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