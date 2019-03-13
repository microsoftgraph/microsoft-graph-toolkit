# Persona Control

## Description
The persona control is used to display a person or contact by using their photo, name, and/or email address. 

## Authentication

The control leverages the global authentication provider described in the [authentication documentation](./../authentication.md) to fetch the required data.

## Graph scopes

This control uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me](https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0) | `User.Read` |
| [/me/photo/$value](https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-beta) | `User.Read` |
| [/me/people/?$search=](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `People.Read` |
| [/users/{id}/photo/$value](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `User.ReadBasic.All` |

> Note: to access the `/me/photo' resource for personal Microsoft Accounts, the beta Graph endpoint is used for all requests

## Example

### Add the control to the html page
```html
<graph-persona person-query=""></graph-persona>
```

## Setting the persona details

The `persona-query` property is used to search the Microsoft Graph for a given person. It will chose the first person available and fetch the persona details. An email works best to ensure the right person is queried, but a name works as well.

Alternatively, the `personaDetails` property is used to set the persona details manually. 

| property | Description |
| --- | --- |
| `personaDetails` | set the user object that will be displayed on the control |

Ex: 

```js
let personaControl = document.getElementById('myPersonaControl');
personaControl.personaDetails = {
    displayName: 'Nikola Metulev',
    email: 'nikola@contoso.com',
    profileImage: 'url'
}
```

## Changing how the control looks

The following attributes are available to customize the behavior

| property  | required  | description |
| --- | --- | --- |
| `view` | optional | enumeration for how the control is layed out - default is `PictureOnly` |
| `image-size` | optional | radius of control - default is 24 |
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