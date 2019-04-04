# Login Control

## Description
A Login Control is a button and flyout control to facilitate login of a user to MSA or AAD. It provides two states:
* When user is not logged in, the control is a simple button to initiate the login process
* When user is logged in, the control displays the current logged in user name, profile image, and email. When clicked, a flyout is opened with a command to logout.

## Authentication

The login control leverages the global authentication provider described in the [authentication documentation](./../authentication.md). 

## Graph scopes

This control uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me](https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0) | `User.Read` |
| [/me/photo/$value](https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-beta) | `User.Read` |

> Note: to access the `/me/photo' resource for personal Microsoft Accounts, the beta Graph endpoint is used for all requests

## Example

### Add the control to the html page
```html
<mgt-login></mgt-login>
```

## Changing how the control looks

The following attributes are available to customize the behavior

[TODO - add all custom properties]

| [css custom properties](../styling-controls.md#css-custom-properties) |
| - |
| `--login-control-background` |


## Using the control without an authentication provider

If you want provide your own logic and authentication, you can use the following properties to set the logged in user's details. 

| property | Description |
| --- | --- |
| `userDetails` | set the user object that will be displayed on the control |

Ex: 

```js
let loginControl = document.getElementById('myLoginControl');
loginControl.userDetails = {
    displayName: 'Nikola Metulev',
    email: 'nikola@contoso.com',
    profileImage: 'url'
}
```

Setting `userDetails` to `null` will go to the logged out state.

Use the `loginInitiated` and `logoutInitiated` events to handle logging in and out. 

## Events

The following events are fired from the control:

| Event | Description |
| --- | --- |
| `loginInitiated` | The user clicked the sign in button to start the login process - cancelable|
| `loginCompleted` | the login process was successful and the user is now signed in |
| `loginFailed` | The user canceled the login process or was unable to log in |
| `logoutInitiated` | The user started to logout - cancelable |
| `logoutCompleted` | the user logged out |
