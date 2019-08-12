---
title: "People-Picker Component"
description: "The mgt-people-picker web component can be used to search and hold a requested amount of people from the Graph."
localization_priority: Normal
author: nmetulev
---

# People-Picker Component

## Description

The `mgt-people-picker` web component can be used to search and hold a requested amount of people from the Graph. By default, it will search for all people, but a group property may be defined for further filtering.

If there are more people to display than the `show-max` value, the search list will not display that person.

## Example

[jsfiddle example](https://jsfiddle.net/metulev/az6pqy2r/)

```html
<mgt-people-picker></mgt-people-picker>
```

![mgt-people](./images/mgt-people-picker.png)

## Properties

By default, the `mgt-people-picker` component fetches events from the `/me/people` endpoint. Use the following attributes to change this behavior:

| property | attribute | Description |
| --- | --- | --- |
| `showMax` | `show-max` | an integer value to indicate the maximum number of people to show - default is 6. |
| `people` | `people` | an array of people to get or set the list of people rendered by the component - use this property to access the people loaded by the component. Set this value to load your own people. |
| `group` | `group` | a string value which belongs to a Graph defined group for further filtering of search.|

Ex:

```html
<mgt-people-picker
  show-max="4">
</mgt-people-picker>
```

## CSS Custom properties

The `mgt-people` component defines these css custom properties

```css
mgt-people-picker {
  --people-list-background-color: blue; /* Background-color for people under search */
  --accent-color: green; /* Color for separator of search input box and people */
}
```
## Graph scopes

This component uses the following Microsoft Graph APIs and permissions:

| resource | permission/scope |
| - | - |
| [/me/people](https://docs.microsoft.com/en-us/graph/api/user-list-people?view=graph-rest-1.0) | `People.Read` |
| [/groups/${groupId}/members](https://docs.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0) | `People.Read` |


## Authentication

The control leverages the global authentication provider described in the [authentication documentation](./../providers.md).
