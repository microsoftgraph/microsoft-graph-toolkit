# People Control

## Description

The `mgt-people` web component can be used to display a group of people or contacts by using their photos or initials. By default, it will display the most frequent contacts for the signed in user.

It utilizes multiple [mgt-person](./person.md) controls, but is able to be bound to a set of people descriptors. If there are more people to display than the `show-max` value, a number will be added to indicate the number of additional contacts.

## Example

```html
<mgt-people></mgt-people>
```

![mgt-people](./images/mgt-people.png)

## Properties

By default, the `mgt-people` component fetches events from the `/me/people` endpoint with the `personType/class eq 'Person'` filter to display frequently contacted users. Use the following attributes to change this behavior:

| attribute | Description |
| --- | --- |
| `show-max` | an integer value to indicate the maximum number of people to show - default is 3. |
| `people` | an array of people to get or set the list of people rendered by the component - use this property to access the people loaded by the component. Set this value to load your own people |

Ex:

```html
<mgt-people
  show-max="4">
</mgt-people>
```

## Custom properties

The `mgt-people` component defines these custom properties

```css
mgt-people {
  --list-margin: 8px 4px 8px 8px; /* Margin for component */
  --avatar-margin: 0 4px 0 0; /* Margin for each person */
}
```

## Templates

The `mgt-people` supports several [templates](../style.md) that allow you to replace certain parts of the component. To specify a template, simply include a `<template>` element inside of a component and set the `data-type` value to one of the following:


### `default` (or when no value is provided)

The default template replaces the entire component with your own. The `people` array is passed to the template as data context

### `no-data`

The template used when no people are available. No data context is passed

### `person`

The template used to render each person. The `person` object is passed to the template as data context

### `overflow`

The template used to render the number beyond the max to the right of the list of people. The `people` array, the `max` number of people currently shown, and the `extra` people beyond are passed to the template as data context

Ex:

```html
<mgt-people>
  <template data-type="person">
    <ul><li data-for="person in people">
      <mgt-person person-query="{{ person.userPrincipalName }}"></mgt-person>
      <h3>{{ person.displayName }}</h3>
      <p>{{ person.jobTitle }}</p>
      <p>{{ person.department }}</p>
    </li></ul>
  </template>
</mgt-people>
```

## Graph scopes

This component uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me/people](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `People.Read` |

## Authentication

The control leverages the global authentication provider described in the [authentication documentation](./../providers.md).
