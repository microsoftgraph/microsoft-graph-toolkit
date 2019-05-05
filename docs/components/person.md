---
title: "Person Component"
description: "The mgt-person web component is used to display a user or contact using their photos or initials."
localization_priority: Normal
author: nmetulev
---

# Person Component

## Description
The person component is used to display a person or contact by using their photo, name, and/or email address. 

## Example

[jsfiddle example](https://jsfiddle.net/metulev/0jkzfr42/)

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

| property  | attribute  | description |
| --- | --- | --- |
| `showName` | `show-name` | set flag to display person display name - default is `false` |
| `showEmail` | `show-email` | set flag to display person email - default is `false` |

## CSS Custom properties

The `mgt-person` component defines these CSS custom properties

```css
mgt-person {
  --avatar-size-s: 24px;
  --avatar-size: 48px; // avatar size when both name and email are shown
  --avatar-font-size--s: 16px;
  --avatar-font-size: 32px; // font-size when both name and email are shown
  --avatar-border: 0;
  --initials-color: white;
  --initials-background-color: magenta;
  --font-size: 12px;
  --font-weight: 500;
  --color: black;
  --email-font-size: 12px;
  --email-color: black;
}
```

[Learn more about styling components](../style.md) 

## Templates

The `mgt-person` component supports several [templates](../templates.md) that allow you to replace certain parts of the component. To specify a template, simply include a `<template>` element inside of a component and set the `data-type` value to one of the following:


### `default` (or when no value is provided)

The default template replaces the entire component with your own. The `person` object is passed to the template as data context

Ex:

```html
<mgt-person>
  <template>
    <div data-if="person.image">
      <img src="{{person.image}}" />
    </div>
    <div data-else>
      {{person.displayName}}
    </div>
  </template>
</mgt-person>
```

## Graph scopes

This control uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me](https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0) | `User.Read` |
| [/me/photo/$value](https://docs.microsoft.com/en-us/graph/api/profilephoto-get?view=graph-rest-beta) | `User.Read` |
| [/me/people/?$search=](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `People.Read` |
| [/me/contacts/*](https://docs.microsoft.com/en-us/graph/api/user-list-contacts?view=graph-rest-1.0&tabs=cs) | `Contacts.Read` |
| [/users/{id}/photo/$value](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `User.ReadBasic.All` |

> Note: to access the `*/photo/$value' resources for personal Microsoft Accounts, the beta Graph endpoint is used

## Authentication

The control leverages the global authentication provider described in the [authentication documentation](./../providers.md) to fetch the required data.