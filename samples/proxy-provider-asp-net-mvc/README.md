# Proxy Provider and ASP.NET MVC

To use the Microsoft Graph Toolkit with backend authentication, one solution is to proxy all calls from the front end to the back end. The Microsoft Graph Toolkit ships with an authentication provider implementation (ProxyProvider) that enables all components to call the Microsoft Graph via the backend. 

This sample is a reference on how to leverage the [ProxyProvider](https://docs.microsoft.com/graph/toolkit/providers/proxy) with an ASP.NET MVC backend. However, it is worth nothing that the ProxyProvider can work with any backend service.

## Client code

The ProxyProvider is instantiated in `Views\Shares\_Layout.cshtml`:

```html
<script src='https://unpkg.com/@Html.Raw("@")microsoft/mgt/dist/bundle/mgt-loader.js'></script>
<script>
    const provider = new mgt.ProxyProvider("https://localhost:44375/api/GraphProxy");
    provider.login = () => window.location.href = '@Url.Action("SignIn", "Account")';
    provider.logout = () => window.location.href = '@Url.Action("SignOut", "Account")';
    mgt.Providers.globalProvider = provider;
</script>
```

This code snippet creates a new ProxyProvider and sets the root api to `/api/GraphProxy`. The root api will be used for every call to the Microsoft Graph initiated by the components - the back end is responsible for translating these calls to actual Microsoft Graph calls and returning the results.

This code snippet also defines a login and logout functions that are used by the Login (`mgt-login`) component to sign in and sign out users.

## Backend code

The `/api/GraphProxy` endpoint that handles all calls to the Microsoft Graph is defined in `Controllers\GraphProxyController.cs`. 

This custom implementation simply proxies every method (**GET**, **POST**, **DELETE**, **PUT**, **PATCH**), and makes the actual call to the Microsoft Graph using the token generated through the Auth2.0 On-Behalf-Of (OBO) flow. Notice that several headers required by the Microsoft Graph are also proxied with each request.

## How to run the completed project

### Prerequisites

To run the completed project in this folder, you need the following:

- [Visual Studio](https://visualstudio.microsoft.com/vs/) installed on your development machine. If you do not have Visual Studio, visit the previous link for download options. (**Note:** This tutorial was written with Visual Studio 2019 version 16.2.3. The steps in this guide may work with other versions, but that has not been tested.)
- Either a personal Microsoft account with a mailbox on Outlook.com, or a Microsoft work or school account.

If you don't have a Microsoft account, there are a couple of options to get a free account:

- You can [sign up for a new personal Microsoft account](https://signup.live.com/signup?wa=wsignin1.0&rpsnv=12&ct=1454618383&rver=6.4.6456.0&wp=MBI_SSL_SHARED&wreply=https://mail.live.com/default.aspx&id=64855&cbcxt=mai&bk=1454618383&uiflavor=web&uaid=b213a65b4fdc484382b6622b3ecaa547&mkt=E-US&lc=1033&lic=1).
- You can [sign up for the Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) to get a free Office 365 subscription.

### Register a web application with the Azure Active Directory admin center

1. Determine your ASP.NET applications's SSL URL. In Visual Studio's Solution Explorer, select the **graph-tutorial** project. In the **Properties** window, find the value of **SSL URL**. Copy this value.


1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.


1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `ASP.NET Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to the ASP.NET app URL you copied in step 1.


1. Choose **Register**. On the **ASP.NET Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.


1. Select **Authentication** under **Manage**. Locate the **Implicit grant** section and enable **ID tokens**. Choose **Save**.


1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and choose **Add**.


1. Copy the client secret value before you leave this page. You will need it in the next step.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now.


### Configure the sample

1. Rename the `PrivateSettings.config.example` file to `PrivateSettings.config`.
1. Edit the `PrivateSettings.config` file and make the following changes.
    1. Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.
    1. Replace `YOUR_APP_PASSWORD_HERE` with the **Application Secret** you got from the App Registration Portal.
1. Open `graph-tutorial.sln` in Visual Studio. In Solution Explorer, right-click the **graph-tutorial** solution and choose **Restore NuGet Packages**.

### Run the sample

In Visual Studio, press **F5** or choose **Debug > Start Debugging**.
