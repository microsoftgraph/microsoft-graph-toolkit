# Login Control

## Description
A Login Control is a button and flyout control to facilitate login of a user to MSA or AAD. It provides two states:
* When user is not logged in, the control is a simple button to initiate the login process
* When user is logged in, the control displays the current logged in user name, profile image, and email. When clicked, a flyout is opened with a command to logout.

## Authentication

The login control leverages the global authentication provider described in the [authentication documentation](./../authentication.md). Make sure you've initialized an authentication provider before the user can use the control

Alternatively, you can initialize the authentication provider directly in HTML without having to set it up in code behind. This can be accomplished by setting the `client-id` directly in HTML.

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
<my-login></my-login>
```

## Changing how the control looks

The following attributes are available to customize the behavior

| property  | required  | description |
| --- | --- | --- |
| `view` | optional | enumeration for how the control is layed out - default is `SmallProfilePhotoRight` |
| `theme` | optional | enumeration for what theme to use - default is `fluent` see [themes](..styling-controls.md#theme/) |
| `custom-style` | optional | use this attribute/property to add additional styles to customize the control. see [styling controls](../styling-controls.md#custom-style/) |


| [css custom properties](../styling-controls.md#css-custom-properties) |
| - |
| `--login-control-background` |

Use the `view` property to control how the control is layed out once the user is logged in:
* PictureOnly - Only show the photo of the user
* EmailOnly - Only show the email of the user
* NameOnly - Only show the display name of the user
* LargeProfilePhotoLeft - show photo, display name, and email - photo on the left
* LargeProfilePhotoRight - show photo, display name, and email - photo on the right
* SmallProfilePhotoLeft - show photo and display name - photo on the left
* SmallProfilePhotoRight - show photo and display name - photo on the right

### Example:
```html
<my-login
    view="pictureOnly"
    theme="fluent"
    custom-style='.login-root-button {background:red;}'>
    </my-login>
```

## Using the control to initialize an authentication provider

The following attributes are available to initialize an authentication provider. The `client-id` attribute must be set for the other properties to have an effect.

| property  | Required  | Description |
| --- | --- | --- |
`client-id` | optional | minimum required property to initialize an authentication provider through the control
`login-type` | optional | either `redirect` or `popup` - default is `redirect`
`scopes` | optional | comma delimited scopes - default `user.read`
`authority` | optional | default is `https://login.microsoftonline.com/common`
`wam` | optional flag | use WAM when running as PWA on Windows

### Example: Initialize an authentication provider
```html
<my-login 
    client-id="client-id">
    </my-login>
```

### Example: Initialize an authentication provider with more options
```html
<my-login 
    client-id="client-id"
    login-type="popup"
    scopes="user.read, calendars.read"
    authority="https://login.microsoftonline.com/contoso.onmicrosoft.com"
    wam>
    </my-login>
```

## Using the control without an authentication provider

If you do not want to use the built in authentication provider, you can use the following properties to set the logged in user's details.

| property  |  Description |
| --- | --- | --- |
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

Setting `userDetails` to `null` or `{}` will go to the logged out state.

Use the `loginInitiated` and `logoutInitiated` events to handle logging in and out. 

## Adding additional commands to dropdown

In addition to the logout command in the dropdown, you can add custom commands that are rendered above the logout command in the dropdown.

| property  |  Description |
| --- | --- | --- |
| `commands` | `LoginDropdownCommand` object that describes each additional command |

`LoginDropdownCommand` is defined as:

```ts
interface LoginDropdownCommand {

    label: string;      // required - text to show
    url?: string;       // optional - url to open in a new window when command is invoked
    action? : function; // optional - callback called when command is invoked
}
```

### Example:
```html
<my-login id="myLoginControl"></my-login>
```

```js
let loginControl = document.getElementById('myLoginControl');
loginControl.commands = [
    {
        label: 'go to profile page',
        url: 'http://myprofilepagelink'
    },
    {
        label: 'settings',
        action: function() {...}
    }
]
```

## Events

The following events are fired from the control:

| Event | Description |
| --- | --- |
| `loginInitiated` | The user started the login process |
| `loginCompleted` | The user completed the login process |
| `loginCanceled` | The user canceled the login process |
| `LogoutInitiated` | The user started to logout |
| `LogoutCompleted` | the user logged out |


## Anatomy
```
div.login-root
|---button.login-root-button
|   |---div.login-header
|       |---div.login-header-user
|       |---div.login-header-user-image-container
|           |---img.login-user-image
|---div.login-dropdown-root

```
