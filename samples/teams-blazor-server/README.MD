# MGT-Teams-Blazor-Server
Small sample Blazor Server application that runs inside a Microsoft Teams Tab and uses the Microsoft Graph Toolkit to login and fetch data from the graph.

![alt text](https://github.com/thomyg/MGT-Teams-Blazor-Server/blob/master/MGT%20Teams%20Blazor%20Server.gif "Demo")

## Short Walkthrough

* Start new sample Blazor Server project in Visual Studio
* Get rid of the left navigation and the top bar, only focus on the main content
* Add script reference to teams sdk and mgt loader in Host.cshtml
* make sure to html encode the @ in URLs you find in other JS examples, Blazor doesn't like that (thx @appieschot!)
* Create Auth.cshtml to be used as auth-pop-url
* Publish application to Azure to get the final URL
* Create Azure AD app for mgt-login config, add Redirect URI to your auth page (used a single tenant app for testing)
* Make sure to enable Implicit grant (Access tokens, ID tokens)
* Add mgt-teams-provider to Index.razor
* Configure mgt-teams-provider with client-id from AAD app, auth-pop-url and authority (needed for single tenant app)
* build and publish your application
* go to Teams, open App Studio and create your application
* add a personal tab and point it to your published application
* install Teams app from App Studio
* try to click on the sign in button

## Pitfalls

* if testing with Teams in browser, make sure to allow pop ups for your auth-pop-url
* you can host your application locally and use ngrok to tunnel your app as well (add/change Azurue AD app redirect URI!)
* if you just run your app locally or open the website without Teams the sign in button does nothing
* you can play with the additional code in Host.cshtml that automates the sign in click, but that's work in progress



