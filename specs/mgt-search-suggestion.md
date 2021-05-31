# mgt-search

The suggestion component based on Graph Suggestion API, it provides a input box and flyout. When a query string typed in the input box, the flyout will rendered by the data from Graph Suggestion API, show some suggestions. it supports 3 kinds of entity types, File/Text/Suggestions, and also opened 3 blank entity types for customizing
The component structure as below
<img src="./images/mgt-search.png" width=400/>

## Supported functionality

| Feature | Priority | Notes |
| ------- | -------- | ----- |
| Retreieve text/people/file/ from Microsoft Graph endpoint based on the query string | P0 | |
| Display the text of the matched query and display the name of the file/people| P0| |
| Display the headImg or create a head icon for suggested people| P0| |
| Display an icon indicating if it's a folder or file and the file type| P0 | Icons needed include generic folder icon, .docx, .pptx, .xlsx, and generic file icon for other file types |
| Display relevant details of the file | P0 | Developer should be able to configure what details are being rendered |
| Provide basic default key down action and mouse action | P0 | when mouse move to an item, the backgroud/cursor should change. when arrow up/down, one item should be selected|
| Provide callback function to get data for supporting users customization actions | P0 | when user click one item, should provide a way for getting the suggested data and support customize actions. when enter key pressed, should provide a way for getting the focused data and support customize actions|

## Templates

 `mgt-search` supports several [templates](../customize-components/templates.md) that you can use to replace certain parts of the component. To specify a template, include a `<template>` element inside a component and set the `data-type` value to one of the following.

| Data type | Data context | Description |
| --- | --- | --- |
| loading | null: no data | The template used to render the state of picker while request to graph is being made. |
| error | null: no data | The template used if user search returns no users. |
| no-data | null: no data | An alternative template used in dropdown list if suggestion returns no any results |
| suggestion-input | null: no data | The template to render search input box. |
| suggested-people-header | null: no data | The template to render people entity label. |
| suggested-query-header | null: no data | The template to render text entity label. |
| suggested-file-header | null: no data | The template to render file entity label. |
| suggested-people | suggestionPeople[]: The people suggestion details list | The template to render people entity. |
| suggested-query | suggestionText[]: The text suggestion details list | The template to render text entity. |
| suggested-blank-{index} | suggestionFile[]: Customizing suggestion details list | The template to render customized entity. index can be 0,1,1 such like suggested-blank-0, developer should inject data and customize the styles by themselves. |

The following examples shows how to use customized template `suggested-query`.

```html
<mgt-search>
    <template data-type="suggested-query">
        <div>{{query}}</div>
    </template>
</mgt-search>


## Proposed Solution

### Example 1: basic usage without any callback function
```<mgt-search></mgt-search>```

### Example 2: get suggestion data and customize the next actions by add a listener
```<mgt-search id="suggestion"> </mgt-search>```
```
#User can get this component then bind callback action to this component 
function getSuggestionValue(suggestionValue) {
    var searchValue = '';
    if (suggestionValue.entity == 'File') {
        searchValue = suggestionValue.name;
    } else if (suggestionValue.entity == 'Text') {
        searchValue = suggestionValue.text;
    } else if (suggestionValue.entity == 'People') {
        searchValue = suggestionValue.displayName;
    }
    return searchValue;
}

document.querySelector('mgt-search').addEventListener('suggestionClick', e => {
    var searchValue = e.detail.displayName;
    window.location.assign('https://www.bing.com/search?q=' + searchValue);
});

document.querySelector('mgt-search').addEventListener('enterPress', e => {
    var originalValue = e.detail.originalValue;
    var suggestedValue = e.detail.suggestedValue;
    var searchValue = getSuggestionValue(suggestedValue);
    window.location.assign('https://www.bing.com/search?q=' + searchValue);
});

```

### Example 3: Developer can select entity types

```<mgt-search suggested-entity-types="file, query, people"></mgt-search>```

```<mgt-search suggested-entity-types="file, people"></mgt-search>```

```<mgt-search suggested-entity-types="people, query, blank-1"></mgt-search>```

### Example 4: Developer can set each entity type's suggestion count

```<mgt-search max-query-suggestions="3" max-file-suggestions="2" max-people-suggestions="3"></mgt-search>```

## Events

The following events are fired from the component.

| Event | Description |
| --- | --- |
| `suggestionClick` | When click an entity , the listener will be trigged. the suggested value of the clicked item as a parameter|
| `enterPress` | When press enter key, the listener will be trigged, it has originalValue( input box value) and suggestedValue (suggestion) as parameters|

## Attributes and Properties

| Attribute | Property | Description |
| --------- | -------- | ----------- |
| `max-query-suggestions` | `maxQuerySuggestions` | The max suggestion count for query. |
| `max-people-suggestions` | `maxPeopleSuggestions` | The max suggestion count for people. |
| `max-file-suggestions` | `maxFileSuggestions` | The max suggestion count for file. |
| `suggested-entity-types` | `selectedEntityTypes` | Suggestion entity types, free combination of text/people/file, use ',' to do the segmentation |
| `other properties` | `other properties` | awaiting for the graph suggestion API onboard. |

## Themes
### light(default)
<img src="./images/mgt-search-light.png" width=400/>

### dark
<img src="./images/mgt-search-dark.png" width=400/>

## CSS custom properties

The `mgt-searach-suggestion` component defines the following CSS custom properties.

```css
mgt-search {

    --suggestion-item-background-color--hover - {Color} background color for an hover item
    --suggestion-list-background-color - {Color} background color
    --suggestion-list-query-color - {Color} Text Suggestion font color
    /* other more properties same with mgt-person / mgt-file */
}
```

## APIs and Permissions

| Query | Use if | Permission Scopes |
| ----- | ------ | ----------------- |
| `awaiting for the graph suggestion API onboard. ` | `awaiting for the graph suggestion API onboard. ` | awaiting for the graph suggestion API onboard.  |

## Extend for more control

We provide a way to get data from default components for development if you don't want to override the component, and also provide a way to override some internal methods.

| Method | Description |
| - | - |
| onClickCallback | include one parameter, suggestion value that user clicked under the dropdown list |
| onEnterKeyPressCallback | include two parameters, the first one is input box value, the second one is suggestion value |
| renderInput | Renders the input text box. |
| renderInput | Renders the input text box. |
| renderFlyout | Renders the flyout chrome. |
| renderFlyoutContent | Renders the appropriate state in the results flyout. |
| renderLoading | Renders the loading state. |
| renderNoData | Renders the state when no results are found for the search query. |
| renderPeopleSearchResults | Renders the list of people search results, include no data processing |
| renderFileSearchResults | Renders the list of file search results. include no data processing |
| renderTextSearchResults | Renders the list of text search results. include no data processing |
| renderSuggestionEntityLabelPeople | Renders the list of people entity label. |
| renderSuggestionEntityLabelText | Renders the list of text entity label. |
| renderSuggestionEntityLabelFile | Renders the list of file entity label. |
| renderSuggestionEntityPeople | Renders the list of people search results, list length > 0, sub method of renderPeopleSearchResults. |
| renderSuggestionEntityText | Renders the list of text search results, list length > 0, sub method of renderTextSearchResults. |
| renderSuggestionEntityFile | Renders the list of file search results, list length > 0, sub method of renderFileleSearchResults. |


