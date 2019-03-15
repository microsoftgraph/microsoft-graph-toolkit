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
<mgt-login
    view="pictureOnly"
    theme="fluent"
    custom-style='.login-root-button {background:red;}'>
    </mgt-login>
```

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

Setting `userDetails` to `null` or `{}` will go to the logged out state.

Use the `loginInitiated` and `logoutInitiated` events to handle logging in and out. 

## Adding additional commands to dropdown

In addition to the logout command in the dropdown, you can add custom commands that are rendered above the logout command in the dropdown.

| property | Description |
| --- | --- |
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
<mgt-login id="myLoginControl"></mgt-login>
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
