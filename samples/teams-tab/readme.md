# Microsoft Teams Tab sample

For the quickest way to run this sample, you will need:

* Visual Studio Code
* [Live Server extension in Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer)
* [ngrok](https://ngrok.com/)

1. Visit the [Microsoft identity platform documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app) to learn more about creating a new application and retrieving a clientId. Add the client id in `index.html`

    > Note: Msal only supports the Implicit Flow for OAuth. Make sure to enable Implicit Flow in your application in the Azure Portal (it is not enabled by default). Under `Authentication`, find the `Implicit grant` section and check both checkboxes for `Access tokens` and `ID tokens`

2. Open `index.html` with Live Server and start `ngrok` in terminal/cmd 

    ```bash
    ngrok http [port]
    ```

    Live Server default port is `8080`

4. Go back to the application registration in the Azure portal and add the redirect url. The redirect url is your ngrok url with the auth page appended.

    Ex: `https://b5d9404b.ngrok.io/auth.html`

5. Now you can use [App Studio](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-app-studio) to quickly develop your app manifest for Microsoft Teams and test the app. Use the ngrok url for the tab url.