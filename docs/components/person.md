# Person Control

## Description
The person control is used to display a person or contact by using their photo, name, and/or email address. 

## Authentication

The control leverages the global authentication provider described in the [authentication documentation](./../providers.md) to fetch the required data.

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

There are three properties that a developer can use to set the person details, use only one of them per instance:

* Set the `user-id` attribute or `userId` property to fetch the user from the Microsoft Graph by using their id.  

* Set the `person-query` attribute or `personQuery` property to search the Microsoft Graph for a given person. It will chose the first person available and fetch the person details. An email works best to ensure the right person is queried, but a name works as well.

* Set the `person-details` attribute or `personDetails` property to manually set the person details.

    Ex: 

    ```js
    let personControl = document.getElementById('myPersonControl');
    personControl.personDetails = {
        displayName: 'Nikola Metulev',
        email: 'nikola@contoso.com',
        image: 'url'
    }
    ```

  If no image is provided, one will be fetched (if available).

## Changing how the control looks

The following attributes are available to customize the behavior

| property  | required  | description |
| --- | --- | --- |
| `showName` | optional | set flag to display person display name - default is `false` |
| `showEmail` | optional | set flag to display person email - default is `false` |


| [css custom properties](../styling-controls.md#css-custom-properties) |
| - |
| `--login-control-background` | 
| TODO
