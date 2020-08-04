# Proxy Provider and ASP.NET Core

To use the Microsoft Graph Toolkit with backend authentication, one solution is to proxy all calls from the front end to the back end. The Microsoft Graph Toolkit ships with an authentication provider implementation (ProxyProvider) that enables all components to call the Microsoft Graph via the backend. 

This sample is a reference on how to leverage the [ProxyProvider](https://docs.microsoft.com/graph/toolkit/providers/proxy) with an ASP.NET Core backend. However, it is worth noting that the ProxyProvider can work with any backend service.

## Client code

The ProxyProvider is instantiated in `Views\Shares\_Layout.cshtml`:

```html
<script src='https://unpkg.com/@Html.Raw("@")microsoft/mgt/dist/bundle/mgt-loader.js'></script>
<script>
    const provider = new mgt.ProxyProvider("/api/proxy");
    provider.login = () => window.location.href = '@Url.Action("SignIn", "Account")';
    provider.logout = () => window.location.href = '@Url.Action("SignOut", "Account")';
    mgt.Providers.globalProvider = provider;
</script>
```

This code snippet creates a new ProxyProvider and sets the root api to `/api/proxy`. The root api will be used for every call to the Microsoft Graph initiated by the components - the back end is responsible for translating these calls to actual Microsoft Graph calls and returning the results.

This code snippet also defines a login and logout functions that are used by the Login (`mgt-login`) component to sign in and sign out users.

## Backend code

The `/api/Proxy` endpoint that handles all calls to the Microsoft Graph is defined in `Controllers\ProxyController.cs`. 

This custom implementation simply proxies every method (**GET**, **POST**, **DELETE**, **PUT**, **PATCH**), and makes the actual call to the Microsoft Graph using the token generated through [MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet). Notice that several headers required by the Microsoft Graph are also proxied with each request.

## How to run the completed project

### Prerequisites

To use the Microsoft Graph Connect Sample for ASP.NET Core 3.1, you need the following:

- Visual Studio 2019 [with .NET Core 3.1 SDK](https://www.microsoft.com/net/download/core) installed on your development computer.
- Either a [personal Microsoft account](https://signup.live.com) or a [work or school account](https://dev.office.com/devprogram). (You don't need to be an administrator of the tenant.)
- The application ID and key from the application that you [register on the App Registration Portal](#register-the-app).

### Register the app

1. Navigate to the [Azure AD Portal](https://portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account** with permissions to create app registrations.

   > **Note:** If you do not have permissions to create app registrations contact your Azure AD domain administrators.

2. Click **Azure Active Directory** from the left-hand navigation menu.

3. Click **App registrations** from the current blade navigation pane.

4. Click **New registration** from the current blade content.

5. On the **Register an application** page, specify the following values:

   - **Name** = [Desired app name]
   - **Supported account types** = [Choose the value that applies to your needs]
   - **Redirect URI**
     - Type (dropdown) = Web
     - Value = `https://localhost:44334/signin-oidc`

   > **Note:** Ensure that the Redirect URI value is unique within your domain. This value can be changed at a later time and does not need to point to a hosted URI. If the example URI above is already used please choose a unique value.

   1. Under **Advanced settings**, set the value of the **Logout URL** to `https://localhost:44334/Account/SignOut`
   2. Copy the **Redirect URI** as you will need it later.

6. Once the app is created, copy the **Application (client) ID** and **Directory (tenant) ID** from the overview page and store it temporarily as you will need both later.

7. Click **Certificates & secrets** from the current blade navigation pane.

   1. Click **New client secret**.
   2. On the **Add a client secret** dialog, specify the following values:

      - **Description** = MyAppSecret1
      - **Expires** = In 1 year

   3. Click **Add**.

   4. After the screen has updated with the newly created client secret copy the **VALUE** of the client secret and store it temporarily as you will need it later.

      > **Important:** This secret string is never shown again, so make sure you copy it now.
      > In production apps you should always use certificates as your application secrets, but for this sample we will use a simple shared secret password.

8. Click **Authentication** from the current blade navigation pane.
   1. Select 'ID tokens'
9. Click **API permissions** from the current blade navigation pane.

   1. Click **Add a permission** from the current blade content.
   2. On the **Request API permissions** panel select **Microsoft Graph**.

   3. Select **Delegated permissions**.
   4. In the "Select permissions" search box type "User".
   5. Select **openid**, **email**, **profile**, **offline_access**, **User.Read**, **User.ReadBasic.All** and **Mail.Send**.

   6. Click **Add permissions** at the bottom of flyout.

   > **Note:** Microsoft recommends that you explicitly list all delegated permissions when registering your app. While the incremental and dynamic consent capabilities of the v2 endpoint make this step optional, failing to do so can negatively impact admin consent.

### Configure and run the sample

1. Download or clone the Microsoft Graph Connect Sample for ASP.NET Core.

2. Open the **MicrosoftGraphAspNetCoreConnectSample.sln** sample file in Visual Studio 2019.

3. In Solution Explorer, open the **appsettings.json** file in the root directory of the project.

   a. For the **AppId** key, replace `ENTER_YOUR_APP_ID` with the application ID of your registered application.

   b. For the **AppSecret** key, replace `ENTER_YOUR_SECRET` with the password of your registered application. Note that in production apps you should always use certificates as your application secrets, but for this sample we will use a simple shared secret password.

4. Press F5 to build and run the sample. This will restore NuGet package dependencies and open the app.

   > If you see any errors while installing packages, make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue.

5. Sign in with your personal (MSA) account or your work or school account and grant the requested permissions.

6. You should see your profile picture and your profile data in JSON on the start page.

7. Change the email address in the box to another valid account's email in the same tenant and choose the **Load data** button. When the operation completes, the profile of the choosen user is displayed on the page.

8. Optionally edit the recipient list, and then choose the **Send email** button. When the mail is sent, a Success message is displayed on the top of the page.

## Key components of the sample

The following files contain code that's related to connecting to Microsoft Graph, loading user data and sending emails.

- [`appsettings.json`](Proxy-Provider-Asp-Net-Core/appsettings.json) Contains values used for authentication and authorization.
- [`Startup.cs`](Proxy-Provider-Asp-Net-Core/Startup.cs) Configures the app and the services it uses, including authentication.

### Controllers

- [`AccountController.cs`](Proxy-Provider-Asp-Net-Core/Controllers/AccountController.cs) Handles sign in and sign out.
- [`HomeController.cs`](Proxy-Provider-Asp-Net-Core/Controllers/HomeController.cs) Handles the requests from the UI.

### Views

- [`Index.cshtml`](Proxy-Provider-Asp-Net-Core/Views/Home/Index.cshtml) Contains the sample's UI.

### Helpers

- [`GraphAuthProvider.cs`](Proxy-Provider-Asp-Net-Core/Helpers/GraphAuthProvider.cs) Gets an access token using MSAL's **AcquireTokenSilent** method.
- [`GraphSdkHelper.cs`](Proxy-Provider-Asp-Net-Core/Helpers/GraphSDKHelper.cs) Initiates the SDK client used to interact with Microsoft Graph.
- [`GraphService.cs`](Proxy-Provider-Asp-Net-Core/Helpers/GraphService.cs) Contains methods that use the **GraphServiceClient** to build and send calls to the Microsoft Graph service and to process the response.
  - The **GetUserJson** action gets the user's profile by an email address and converts it to JSON.
  - The **GetPictureBase64** action gets the user's profile picture and converts it to a base64 string.
  - The **SendEmail** action sends an email on behalf of the current user.
