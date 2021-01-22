# mgt-file-picker

The file picker component provides a button, a drop down list of files, and as well as button to launch a full file browsing experience to enable a user to select a file from OneDrive or SharePoint for further action. This component uses the [mgt-file-list](./mgt-file-list.md) component to display files in the dropdown. The selected file object(s) is returned to the developer.

<img src="./images/mgt-file-picker.png"/>

## Supported functionality

| Feature | Priority | Notes |
| ------- | -------- | ----- |
| Provide a button that renders a drop down when clicked | P0 | |	
| Retrieve a list of files based on information provided | P0 | |
| Use the `mgt-file-list` component to render the list inside of the drop down | P0 | |
| On selection or deselection of a file, an event is fired | P0 | |	
| The selected file object(s) is accessible by the developer for further action |P0 | |	
| Ability for developer to toggle between single-select and multiselect | P1 |  | 
| Visual indication of selected files | P2 | Will not be included in v1. |
| Render the selected files | P2 | Will not be included in v1. We will provide a sample for developers to render selected files using `mgt-file` components themselves. |
| Extended file picker is launched when “See all” is clicked | P2 | Not included in v1. |

## Proposed Solution

### Example 1: Show the user's top 10 most recently used files to select from
```<mgt-file-picker insight-type=“used” show-max=“10”></mgt-file-picker>```

### Example 2: Show a specific folder in OneDrive to select files from
```<mgt-file-picker drive-id=“123” item-id=“456”></mgt-file-picker>```

## Attributes and Properties

| Attribute | Property | Description |
| --------- | -------- | ----------- |
| `file-list-query` | `fileListQuery` | The full query or path to the folder in a drive or site to retrieve the list of files from |
| `file-queries` | `fileQueries` | An array of file queries to be rendered by the component |
| `files` | `files` | An array of files to get or set the list of files rendered by the component. Use this to access the files loaded by the component. Set this value to load your own files. |
| `insight-type` | `insightType` | Set to show the user’s trending, used, or shared files. |
| `drive-id` | `driveId` | Id of the drive the folder belongs to. Must also provide either `item-id` or `item-path`. |
| `group-id ` | `groupId` | Id of the group the folder belongs to. Must also provide either `item-id` or `item-path`. |
| `site-id` | `siteId` | Id of the site the folder belongs to. Must also provide either `{item-id}` or `{item-path}`. Provide `{list-id}` if you’re referencing a file from a list. |
| `item-id` | `itemId` | Id of the folder. Default query is `/me/drive/items`. Provide `{drive-id}`, `{group-id}`, `{site-id}`, or `{user-id}` to query a specific location. |
| `item-path` | `itemPath` | Item path of the folder. Default query is `/me/drive/root`. Provide `{drive-id}`, `{group-id}`, `{site-id}`, or `{user-id}` to query a specific location. |
| `show-max` | `showMax` | A number value to indicate the maximum number of files to show. |
| `selected-files` | `selectedFiles`| An array of selected file objects. |
| `multiselect` | `multiselect` | Enable/Disable selection of more than 1 file. Default is `false`.|

## Events

| Event | Description |
| ----- | ----------- |
|selectionChanged | Fired when the user selects or deselects a file |

## APIs and Permission

Same as [mgt-file-list](./mgt-file-list.md).

## Templates

| Data type | Data Context | Description |
| --------- | ------------ | ----------- |
| default | null: no data | The template used to override the rendering of the entire component |
| loading | null: no data | The template used to render the state of the picker while the request to Graph is being made. |
| error | null: no data | The template used if no files are returned |
| selected-file | file: the file details object | The template to render selected files |
| file | file: the file details object | The template to render files in the dropdown |