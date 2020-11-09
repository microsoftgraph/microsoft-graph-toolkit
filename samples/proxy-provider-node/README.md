# Node.js Graph Proxy sample

This sample implements a Node.js server that acts as a proxy for Microsoft Graph. This can be used along with the Microsoft Graph Toolkit's [proxy provider](https://docs.microsoft.com/graph/toolkit/providers/proxy) to make all Microsoft Graph API calls from a backend service.

## Authorization modes

The proxy service supports to authorization modes: pass-through and on-behalf-of.

> **NOTE**
> The [sample client application](./client) requires using the on-behalf-of mode.

### Pass-through

In pass-through mode, the service takes whatever bearer token is sent in the `Authorization` header from the client and tries to use that to call Microsoft Graph. This mode is primarily for ease of testing.

### On-behalf-of

In on-behalf-of mode, the service uses the Microsoft identity platform's [on-behalf-of flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) to exchange the token sent by the calling client for a Microsoft Graph token. The next section gives instructions on how to register the app in Azure Active Directory.

#### App registration

This sample requires two app registrations. This is needed to take advantage of [combined consent](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#default-and-combined-consent) with the on-behalf-of flow in the Microsoft identity platform. Create an app registration for:

- The Graph Proxy service
- The React front-end app

##### Graph Proxy service

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Node.js Graph Proxy`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, leave the value empty.

1. Select **Register**. On the **Node.js Graph Proxy** page, copy the value of the **Application (client) ID** and **Directory (tenant) ID**. Save them, you will need them in the next step.

1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and select **Add**.

1. Copy the client secret value before you leave this page. You will need it in the next step.

    > **IMPORTANT**
    > This client secret is never shown again, so make sure you copy it now.

1. Select **API permissions** under **Manage**, then select **Add a permission**.

1. Select **Microsoft Graph**, then **Delegated permissions**.

1. Select the following permissions, then select **Add permissions**.

    - **Presence.Read** - this will allow the app to read the authenticated user's presence.
    - **Tasks.ReadWrite** - this allows the app to read and write the user's To-Do tasks.
    - **User.Read** - this will allow the app to read the user's profile and photo.

    > **NOTE**
    > These are the permissions required for the Microsoft Graph Toolkit components used in the client application. If you use different components, you may require additional permissions. See the [documentation](https://docs.microsoft.com/graph/toolkit/overview) for each component for details on required permissions.

1. Select **Expose an API**. Select the **Set** link next to **Application ID URI**. Accept the default and select **Save**.

1. In the **Scopes defined by this API** section, select **Add a scope**. Fill in the fields as follows and select **Add scope**.

    - **Scope name:** `access_as_user`
    - **Who can consent?: Admins and users**
    - **Admin consent display name:** `Access Graph Proxy as the user`
    - **Admin consent description:** `Allows the app to call Microsoft Graph through a proxy service on users' behalf.`
    - **User consent display name:** `Access Graph Proxy as you`
    - **User consent description:** `Allows the app to call Microsoft Graph through a proxy service as you.`
    - **State: Enabled**

1. Copy the new scope. You'll need this later.

##### React front-end app

1. Return to **App registration** in the Azure portal, then select **New registration**.

1. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Node.js Graph Proxy Client`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `http://localhost:8000/authcomplete`.

1. Select **Register**. On the **Node.js Graph  Client** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

1. Select **API permissions** under **Manage**, then select **Add a permission**.

1. Select **APIs my organization uses**, then search for `Node.js Graph Proxy`. Select **Node.js Graph Proxy** in the list.

1. Select the **access_as_user** permission, then select **Add permissions**.

1. In the **Configured permissions** list, remove the **User.Read** permission under **Microsoft Graph**.

##### Add client to proxy's known applications

1. Return to the **Node.js Graph Proxy** app registration in the Azure portal, then select **Manifest** under **Manage**.

1. Locate the `"knownClientApplications": [],` line and replace it with the following, where `CLIENT_APP_ID` is the application ID of the **Node.js Graph Proxy Client** registration.

    ```json
    "knownClientApplications": ["CLIENT_APP_ID"],
    ```

1. Select **Save**.

## Configuring the sample

1. Rename the example.env file to .env, and set the values as follows.

    | Setting | Value |
    |---------|-------|
    | GRAPH_HOST | Set to the specific [Graph endpoint](https://docs.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints) for your organization. |
    | AUTH_PASS_THROUGH | Set to `true` to enable pass-through mode, `false` for on-behalf-of mode. |
    | PROXY_APP_ID | The application ID of your **Node.js Graph Proxy** app registration |
    | PROXY_APP_TENANT_ID | The tenant ID from your **Node.js Graph Proxy** app registration |
    | PROXY_APP_SECRET | The client secret from your **Node.js Graph Proxy** app registration |

1. Rename the ./client/config.example.js to config.js and set the values as follows.

    - Replace `YOUR_PROXY_CLIENT_APP_ID` with the application ID of your **Node.js Graph Proxy Client** app registration.
    - Replace `YOUR_PROXY_APP_ID` with the application ID of your **Node.js Graph Proxy** app registration.

## Run the sample

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

1. Open your browser and go to `http://localhost:8000`.

1. Select the **Sign In** button to sign in. After signing in and granting consent, the app should load data into the components.

### How do I know it worked?

If everything worked correctly, the components should load data just as they would normally without using the proxy provider. You can verify that requests are going through the proxy (rather than directly to Graph) by reviewing the console output in your CLI where you ran `yarn start`.

```Shell
⚡️[server]: Server is running at http://localhost:8000
Auth mode: on-behalf-of
GET /v1.0/me
Accept: */*
POST /v1.0/$batch
Accept: */*
Content-Type: application/json
GET /v1.0/me
Accept: */*
POST /v1.0/$batch
Accept: */*
Content-Type: application/json
GET /beta/me/outlook/taskGroups
Accept: */*
```
