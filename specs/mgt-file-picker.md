# mgt-file-picker

The file picker component provides a button, a drop down list of files, and as well as button to launch a full file browsing experience to enable a user to select a file from OneDrive or SharePoint for further action. This component uses the [mgt-file](./mgt-file.md) components to display files in the dropdown. The selected file object(s) is returned to the developer.

<img src="./images/mgt-file-picker.png"/>

## Supported functionality

| Feature | Priority | Notes |
| ------- | -------- | ----- |
| Provide a button that renders a drop down when clicked | P0 | |	
| Retrieve files based on specified insight type | P0 | |
| Use `mgt-file` components to render each file inside of the drop down | P0 | |
| On selection of a file, a `selectionChanged` event is fired | P0 | |	
| The selected file object(s) is accessible by the developer for further action |P0 | |	
| Extended file picker is launched when “See all” is clicked | P0 | |
| Ability for developer to toggle between single-select and multiselect | P1 | Will not be included in v1. | 
| Visual indication of selected files | P2 | Will not be included in v1. |
| Render the selected files | P2 | Will not be included in v1. We will provide a sample for developers to render selected files using `mgt-file` components themselves. |

## Proposed Solution

### Example 1: Show the user's top 10 most recently used files to select from
```<mgt-file-picker insight-type=“used” show-max=“10”></mgt-file-picker>```

## Attributes and Properties

| Attribute | Property | Description |
| --------- | -------- | ----------- |
| `files` | `files` | An array of files to get or set the list of files rendered by the component. Use this to access the files loaded by the component. Set this value to load your own files. |
| `insight-type` | `insightType` | Set to show the user’s trending, used, or shared files. |
| `show-max` | `showMax` | A number value to indicate the maximum number of files to show. |

## Events

| Event | Description |
| ----- | ----------- |
|selectionChanged | Fired when the user selects a file |

## APIs and Permission

| Query | Use if | Permission Scopes |
| ----- | ------ | ----------------- |
| `GET /me/insights/trending` | `insight-type` is trending | Sites.Read.All |
| `GET /me/insights/used` | `insight-type` is `used` | " |
| `GET /me/insights/shared` | `insight-type` is shared | " |

## Templates

| Data type | Data Context | Description |
| --------- | ------------ | ----------- |
| default | null: no data | The template used to override the rendering of the entire component |
| loading | null: no data | The template used to render the state of the picker while the request to Graph is being made. |
| error | null: no data | The template used if no files are returned |
| file | file: the file details object | The template to render files in the dropdown |