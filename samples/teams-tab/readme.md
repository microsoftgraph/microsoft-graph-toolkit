# Microsoft Teams Tab sample

For the quickest way to run this sample, you will need:

* Visual Studio Code
* [Live Server extension in Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer)
* [ngrok](https://ngrok.com/)

1. Visit the [Microsoft identity platform documentation](https://learn.microsoft.com/azure/active-directory/develop/quickstart-register-app) to learn more about creating a new application and retrieving a clientId. Add the client id in `index.html`

    > Note: Msal only supports the Implicit Flow for OAuth. Make sure to enable Implicit Flow in your application in the Azure Portal (it is not enabled by default). Under `Authentication`, find the `Implicit grant` section and check both checkboxes for `Access tokens` and `ID tokens`

2. Open `index.html` with Live Server and start `ngrok` in terminal/cmd 

    ```bash
    ngrok http [port]
    ```

    Live Server default port is `8080`

4. Go back to the application registration in the Azure portal and add the redirect url. The redirect url is your ngrok url with the auth page appended.

    Ex: `https://b5d9404b.ngrok.io/auth.html`

5. Now you can use the [Developer Portal for Teams](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/teams-developer-portal) to configure, distribute and manage your application. You can access the [Developer Portal for Teams in a web browser](https://dev.teams.microsoft.com/) or as a [Teams App](https://teams.microsoft.com/l/app/14072831-8a2a-4f76-9294-057bf0b42a68). You can also use the [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) in **Visual Studio Code** to quickly create and deploy your Teams app.