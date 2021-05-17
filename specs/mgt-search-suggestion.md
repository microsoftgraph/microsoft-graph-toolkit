# mgt-search-suggestion

The suggestion component based on Graph Suggestion API, it provides a input box and flyout. When a query string typed in the input box, the flyout will rendered by the data from Graph Suggestion API, show some suggestions. it supports 3 kinds of entity types, File/Text/Suggestions. The component structure as below
<img src="./images/mgt-search-suggestion.png" width=400/>

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

 `mgt-search-suggestion` supports several [templates](../customize-components/templates.md) that you can use to replace certain parts of the component. To specify a template, include a `<template>` element inside a component and set the `data-type` value to one of the following.

| Data type | Data context | Description |
| --- | --- | --- |
| default | null: no data | The template used to override the rendering of the entire component.
| loading | null: no data | The template used to render the state of picker while request to graph is being made. |
| error | null: no data | The template used if user search returns no users. |
| no-data | null: no data | An alternative template used if user search returns no users. |
| search-suggestion-input | null: no data | The template to render search input box. |
| search-suggestion-label-people | null: no data | The template to render people entity label. |
| search-suggestion-label-text | null: no data | The template to render text entity label. |
| search-suggestion-label-file | null: no data | The template to render file entity label. |
| search-suggestion-people | suggestionPeople[]: The people suggestion details list | The template to render people entity. |
| search-suggestion-text | suggestionText[]: The text suggestion details list | The template to render text entity. |
| search-suggestion-file | suggestionFile[]: The file suggestion details list | The template to render file entity. |

The following examples shows how to use the `error` template.

```html
<mgt-search-suggestion>
    <template data-type="search-suggestion-texts">
        <div  data-for="text in texts">
        {{text.text}}
        </div>
    </template>
</mgt-search-suggestion>

```

## Proposed Solution

### Example 1: basic usage without any callback function
```<mgt-search-suggestion></mgt-search-suggestion>```

### Example 2: Developer provides a site-id and item-id
```<mgt-search-suggestion id="search-suggestion"> </mgt-search-suggestion>```
```
#User can get this component then bind callback action to this component 

function onClickCallback(suggestionValue) {
    console.log('suggestion value:', suggestionValue);
    var searchValue = getSuggestionValue(suggestionValue);
    window.location.assign('https://www.bing.com/search?q=' + searchValue);
}

function onEnterKeyPressCallback(originalValue, selectedSuggestionValue) {
    console.log('original value:', originalValue);
    console.log('suggestion value:', selectedSuggestionValue);
    var searchValue = getSuggestionValue(selectedSuggestionValue);
    window.location.assign('https://www.bing.com/search?q=' + searchValue);
}

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

var obj = document.getElementById('search-suggestion');
obj.onClickCallback = onClickCallback;
obj.onEnterKeyPressCallback = onEnterKeyPressCallback;

```

### Example 3: Developer can select entity types

```<mgt-search-suggestion selected-entity-types="file, text, people"></mgt-search-suggestion>```

```<mgt-search-suggestion selected-entity-types="file, people"></mgt-search-suggestion>```

```<mgt-search-suggestion selected-entity-types="people, text"></mgt-search-suggestion>```

## Example 4: Developer can set each entity type's suggestion count

```<mgt-search-suggestion max-text-suggestion-count="3" max-file-suggestion-count="2" max-people-suggestion-count="3"></mgt-search-suggestion>```

## Attributes and Properties

| Attribute | Property | Description |
| --------- | -------- | ----------- |
| `max-text-suggestion-count` | `maxTextSuggestionCount` | The max suggestion count for text. |
| `max-people-suggestion-count` | `maxPeopleSuggestionCount` | The max suggestion count for people. |
| `max-file-suggestion-count` | `maxFileSuggestionCount` | The max suggestion count for file. |
| `selected-entity-types` | `selectedEntityTypes` | Suggestion entity types, free combination of text/people/file, use ',' to do the segmentation |
| `other properties` | `other properties` | awaiting for the graph suggestion API onboard. |

## Themes
### light(default)
<img src="./images/mgt-search-suggestion-light.png" width=400/>

### dark
<img src="./images/mgt-search-suggestion-dark.png" width=400/>

## CSS custom properties

The `mgt-searach-suggestion` component defines the following CSS custom properties.

```css
mgt-search-suggestion {

    --suggestion-item-background-color--hover - {Color} background color for an hover item
    --suggestion-list-background-color - {Color} background color
    --suggestion-list-text-color - {Color} Text Suggestion font color
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


