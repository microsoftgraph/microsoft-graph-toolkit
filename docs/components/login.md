---
title: "Login Component"
description: "mgt-login displays current signed in user. "
localization_priority: Normal
author: nmetulev
---

# Login Component

## Description
A Login Component is a button and flyout control to facilitate Microsoft Identity authentication. It provides two states:
* When user is not logged in, the control is a simple button to initiate the login process
* When user is logged in, the control displays the current logged in user name, profile image, and email. When clicked, a flyout is opened with a command to logout.

## Example

[jsfiddle example](https://jsfiddle.net/metulev/scb9muh4)

```html
<mgt-login></mgt-login>
```

## Using the control without an authentication provider

The component works with a provider and the Microsoft Graph out of the box. However, if you want to provide your own logic and authentication, you can use the following properties to set the logged in user's details. 

| property | attribute | Description |
| --- | --- | -- |
| `userDetails` | `user-details` | set the user object that will be displayed on the control |

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

## CSS Custom properties

The `mgt-login` component defines these CSS custom properties

```css
mgt-login {
  --font-size: 14px;
  --font-weight: 600;
  --width: '100%';
  --height: '100%';
  --margin: 0;
  --padding: 12px 20px;
  --color: #201f1e;
  --background-color: transparent;
  --background-color--hover: #edebe9;
  --popup-content-background-color: white;
  --popup-command-font-size: 12px;
  --popup-command-margin: 0;
  --popup-padding: 16px;
}
```

[Learn more about styling components](../style.md) 

## Events

The following events are fired from the control:

| Event | Description |
| --- | --- |
| `loginInitiated` | The user clicked the sign in button to start the login process - cancelable|
| `loginCompleted` | the login process was successful and the user is now signed in |
| `loginFailed` | The user canceled the login process or was unable to log in |
| `logoutInitiated` | The user started to logout - cancelable |
| `logoutCompleted` | the user logged out |

## Graph scopes

This component leverages the [Person component](./person.md) for displaying the user and inherits all scopes. 

## Authentication

The login control leverages the global authentication provider described in the [authentication documentation](./../providers.md). 