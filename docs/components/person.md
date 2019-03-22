# Person Control

## Description
The person control is used to display a person or contact by using their photo, name, and/or email address. 

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
<mgt-person person-query=""></mgt-person>
```

## Setting the person details

The `person-query` property is used to search the Microsoft Graph for a given person. It will chose the first person available and fetch the person details. An email works best to ensure the right person is queried, but a name works as well.

Alternatively, the `personDetails` property is used to set the person details manually. 

| property | Description |
| --- | --- |
| `personDetails` | set the user object that will be displayed on the control |

Ex: 

```js
let personControl = document.getElementById('myPersonControl');
personControl.personDetails = {
    displayName: 'Nikola Metulev',
    email: 'nikola@contoso.com',
    image: 'url'
}
```

## Changing how the control looks

The following attributes are available to customize the behavior

| property  | required  | description |
| --- | --- | --- |
| `showName` | optional | set flag to display person display name - default is `false` |
| `showEmail` | optional | set flag to display person email - default is `false` |
| `image-size` | optional | radius of control - default is 24 |


| [css custom properties](../styling-controls.md#css-custom-properties) |
| - |
| `--login-control-background` |
