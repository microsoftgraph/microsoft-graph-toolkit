# mgt-taxonomy-picker

## Overview
The taxonomy picker is a component that can query the Microsoft Graph API for Taxonomy and render a dropdown control allowing selection of **a single** term based on the specified term set id and a combination of the specified term set id and the specified term id. 

> **Note**
> This capability used to support a single selection coming from the term store. This will be reviewed in the future.  

## User Scenarios

Let's assume a term store with the following structure:

![mgt-taxonomy-pikcer-structure](./images/mgt-taxonomy-pikcer-structure.png)

### Scenario 1 - Select a term from a term set

Shows the terms marked with a red box in the image above.

![mgt-taxonomy-pikcer-termset](./images/mgt-taxonomy-pikcer-termset.png)

### Scenario 2 - Select a term from a term in a term set

Shows the terms marked with a green box in the image above.

![mgt-taxonomy-pikcer-term](./images/mgt-taxonomy-pikcer-term.png)


### Scenario 3 - Select a term from a term in a term set in a different language

Shows French labels of the terms marked with a green box in the image above.

![mgt-taxonomy-pikcer-locale-term](./images/mgt-taxonomy-pikcer-locale-term.png)


## Proposed Solution

### Examples

#### Show the first level children of the term set with the specified id.

`<mgt-taxonomy-picker termset-id="138a652e-7f23-46f6-b480-13da2308c235"></mgt-taxonomy-picker>`

#### Show the first level children of the term with the specified id under the term set with the specified id.

`<mgt-taxonomy-picker termset-id="138a652e-7f23-46f6-b480-13da2308c235" term-id="a56caeb7-3b7d-4d22-93a9-0232e12905f6"></mgt-taxonomy-picker>`

#### Show the first level children of the term with the specified id under the term set with the specified id in a different language.

`<mgt-taxonomy-picker termset-id="138a652e-7f23-46f6-b480-13da2308c235" term-id="a56caeb7-3b7d-4d22-93a9-0232e12905f6" locale="fr-FR"></mgt-taxonomy-picker>`

#### Cache the response from the API call for 50 seconds.

`<mgt-taxonomy-picker termset-id="138a652e-7f23-46f6-b480-13da2308c235" cache-enabled="true" cache-invalidation-period="50000"></mgt-taxonomy-picker>`

## Properties and Attributes

| Attribute                 | Property                | Description                                                                                                                                                                                                                           | Type                       |
| ------------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------- |
| placeholder                  | placeholder                 | The placeholder to use in the combobox. No default value is provided.                                                                                                                                        | string                     |
| termset-id                  | termsetId                 | The Id of the termset where the terms are present. The terms under the termset will be shown if `term-id` is not passed                                                                                                                                         | string                     |
| term-id                  | termId                 | The id of the term where the terms are present.                                                                                                                                            | Number                     |
| locale                  | locale                | The locale of the terms that need to be displayed. This will be useful only when terms have multiple lables in different languages.                                                                                                                                                                        | String                     |
| default-selected-term-id                    | defaultSelectedTermId                  | Optional. The id of the term that should be selected by default.     | String |
| version                   | version                 | Optional API version to use when making the GET request. Default is `beta`.                                                                                                                                                           | String                     |
| cache-enabled             | cacheEnabled            | Optional Boolean. When set, it indicates that the response from the resource will be cached. Default is `false`.                                                                                  | Boolean                    |
| cache-invalidation-period | cacheInvalidationPeriod | Optional number of milliseconds. When set in combination with `cacheEnabled`, the delay before the cache reaches its invalidation period will be modified by this value. Default is `0` and will use the default invalidation period. | Number                     |



## Events

| Event            | When is it fired                                | Custom data |
| ---------------- | ----------------------------------------------- | ----------- |
| selectionChanged | Fired when the user makes a change in selection | The selected term which will of the type `TermStore.Term`        |

## Templates

| Data type     | Data Context                           | Description                                                                                  |
| ------------- | -------------------------------------- | -------------------------------------------------------------------------------------------- |
| default       | null: no data                          | The template used to override the rendering of the entire component.                         |
| loading       | null: no data                          | The template used to render the state of the picker while the request to Graph is being made. |
| error         | null: no data                          | The template used there is an error.
| no-data         | null: no data                     | The template used if no terms are present.                         |                                              |

## APIs and Permissions

Permissions required by this component are `TermStore.Read.All`.