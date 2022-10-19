# mgt-picker

## Overview
The picker is a component that can query any Microsoft Graph API and render a dropdown control allowing selection of **a single** resource. The picker will also support a predefined list of resources that won't require to specify a resource from Microsoft Graph.

> **Note**
> This capability used to support a single selection coming from multiple resources. This will be reviewed in the future.  

## User Scenarios

### Select a To Do list (P0)

**Resources:** To Do

<img src="./images/mgt-picker.png"/>


### Select any resource (P0)

**Resources:** Any endpoint on Microsoft Graph delivering an array of resources 

## Proposed Solution

### Example: To Do
`<mgt-picker entity-type="task-lists"></mgt-picker>`

### Example: Generic endpoint
`<mgt-picker resource="/groups"></mgt-picker>`

`<mgt-picker resource="/groups" version="beta"></mgt-picker>`

`<mgt-picker resource="/groups" version="beta" scopes="Group.Read.All"></mgt-picker>`

`<mgt-picker resource="/groups" version="beta" scopes="Group.Read.All" max-pages="5"></mgt-picker>`

`<mgt-picker resource="/groups" version="beta" scopes="Group.Read.All" cache-enabled="true" cache-invalidation-period="50000"></mgt-picker>`

## Properties and Attributes

| Attribute                 | Property                | Description                                                                                                                                                                                                                           | Type                       |
| ------------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------- |
| show-max                  | showMax                 | Specify the number of results to show for the resource. Not in use if `max-pages` is used.                                                                                                                                            | Number                     |
| entity-type               | entityType              | Specify the entity to be fetched and rendered in the picker                                                                                                                                                                           | `todo-lists`               |
| resource                  | resource                | The resource to get from Microsoft Graph (for example, `/me`).                                                                                                                                                                        | String                     |
| scopes                    | scopes                  | Optional array of strings if using the property or a comma delimited scope if using the attribute. The component will use these scopes (with a supported provider) to ensure that the user has consented to the right permission.     | String or array of strings |
| version                   | version                 | Optional API version to use when making the GET request. Default is `v1.0`.                                                                                                                                                           | String                     |
| max-pages                 | maxPages                | Optional number of pages (for resources that support paging). Default is 3. Setting this value to 0 will get all pages.                                                                                                               | Number                     |
| cache-enabled             | cacheEnabled            | Optional Boolean. When set, it indicates that the response from the resource will be cached. Overriden if `refresh()` is called. Default is `false`.                                                                                  | Boolean                    |
| cache-invalidation-period | cacheInvalidationPeriod | Optional number of milliseconds. When set in combination with `cacheEnabled`, the delay before the cache reaches its invalidation period will be modified by this value. Default is `0` and will use the default invalidation period. | Number                     |


## Methods

| Method                  | Description                                                                                                                                 |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------- |
| refresh(force?:boolean) | Call the method to refresh the data. By default, the UI will only update if the data changes. Pass `true` to force the component to update. |

## Events

| Event            | When it is fired                                |
| ---------------- | ----------------------------------------------- |
| selectionChanged | Fired when the user makes a change in selection |

## Templates

| Data type     | Data Context                           | Description                                                                                  |
| ------------- | -------------------------------------- | -------------------------------------------------------------------------------------------- |
| default       | null: no data                          | The template used to override the rendering of the entire component.                         |
| loading       | null: no data                          | The template used to render the state of the picker while the request to Graph is being made |
| error         | null: no data                          | The template used if search returns no results.                                              |
| input-text    | null: no data                          | The template used to render the default text inside of the input box.                        |
| rendered-item | renderedItem: the item being rendered  | The template used to render the item inside the dropdown.                                    |
| selected-item | selectedItem: the selected item object | The template used to render the selected item inside the input box.                          |



## APIs and Permissions

| API                                                                                                         | Use When                  | Permissions |
| ----------------------------------------------------------------------------------------------------------- | ------------------------- | ----------- |
| GET [/me/tasks/lists](https://learn.microsoft.com/graph/api/tasks-list-lists?view=graph-rest-beta&tabs=http) | entity-type is task-lists | Tasks.Read  |

Permissions required by this component when using the `resource` property depend on the data that you want to retrieve with it from Microsoft Graph. For more information about permissions, see the Microsoft Graph [permissions reference](https://learn.microsoft.com/graph/permissions-reference).