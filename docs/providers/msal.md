# MSAL provider

The MSAL Provider uses [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js) to sign in users and acquire tokens to use with the Graph

## Getting started

You can initialize the provider in HTML or JavaScript.

### In your HTML page
Initialize the provider inside your html using the `mgt-msal-provider` component. This will create a new `UserAgentApplication` instance that will be used for all authentication and acquiring tokens.

```html
<mgt-msal-provider client-id="<YOUR_CLIENT_ID>"
                   login-type="redirect/popup"
                   authority=""></mgt-teams-provider>
```
>NOTE: `login-type` and `authority` are optional.

### In JavaScript

Initializing the provider in JavaScript allows you to provide more options or use an existing `UserAgentApplication`

To initialize a provider by creating a new `UserAgentApplication`, you can pass a MsalConfig object to the constructor and assign the new object to the `Provider.globalProvider` variable;

```js
//todo
```

To initialize a provider with an existing 'UserAgentApplication'...

## Creating an app and Client Id

