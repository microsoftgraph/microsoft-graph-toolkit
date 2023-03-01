# mgt-conversation

## Overview
The conversation component renders the content of a 1:1 and a group chat. It leverages the Teams Microsoft Graph APIs to get its data and is using the [Azure Communication Services UI Library](https://azure.github.io/communication-ui-library/) as the main renderer of the conversation.

> **Note**
> This component will be a **React-Only** component in its first implementation. The hard dependency on the ACS UI Library doesn't allow this component to be available as a web component. Another transient dependency will be [Fluent React v8](https://developer.microsoft.com/en-us/fluentui#/controls/web) compared to the rest of the Toolkit that is built on [Fluent Web Components](https://learn.microsoft.com/en-us/fluent-ui/web-components/).

## User Scenarios (in scope)

| Scenario                                                                                                         | Priority | Comments                                                                                                   |
|------------------------------------------------------------------------------------------------------------------|----------|------------------------------------------------------------------------------------------------------------|
| As a user, I want to chat with a specific user                                                                   | P1       |                                                                                                            |
| As a user, I want to chat with a group of users                                                                  | P1       |                                                                                                            |
| As a user, I want to read the complete thread from the chat                                                      | P1       | When loading a chat, the previous messages are shown                                                       |
| As a user, I want to send messages to the other chat participants                                                | P1       |                                                                                                            |
| As a user, I want to receive messages from a chat thread instantly                                               | P1       | Using a websocket that gets the data back from Microsoft Graph                                             |
| As a user, I want to see the avatar (`mgt-person`) and the card (`mgt-person-card`) of people in the chat thread | P1       | The avatar controls from ACS should be replaced by MGT person-specific controls                            |
| As a user, I want to be able to edit messages I sent                                                             | P1       | Changes should be applied instantly for other participants                                                 |
| As a user, I want to be able to delete messages I sent                                                           | P1       | Messages should be removed automatically and a system message should be shown that the message was deleted |
| As a user, I want to see a system message when a participant is added to a group discussion                      | P1       |                                                                                                            |
| As a user, I want to see a system message when a participant is removed from a group discussion                  | P1       |                                                                                                            |
| As a user, I want to see the list of participants in a chat thread                                               | P1       | Relies on a Preview feature of ACS UI Library                                                              |
| As a user, I want to see if my message was seen in 1:1 chat threads                                              | P1       | A visual indication should be visible                                                                      |
| As a user, I want to be able to mention participants in a chat thread                                            | P1       | @ trigger should invoke the `mgt-people-picker` to help with searching users                               |
| As a user, I want to be able to format my messages using a rich text editor                                      | P1       | Depends on ACS UI Library features                                                                         |
| As a user, I want to be able to add emojis to the messages I send                                                | P1       | Leveraging the OOTB Win+. should be enough                                                                 |
| As a user, I want to be able to jump to the full Teams experience from my conversation                           | P1       | Deep linking in Teams                                                                 |
| As a user, I want to be able to react to messages                                                                | P2       | Depends on ACS UI Library features                                                                         |
| As a user, I want to share files with the participants of the chat thread                                        | P2       | Depends on ACS UI Library features and should integrate with our `mgt-file-list` component                 |
| As a user, I want to set the importance of a message when sending it                                             | P2       | Currently not supported by ACS UI Library                                                                  |
| As a user, I want to see a typing indicator when other people are writing messages                               | P3       | Blocked by lack of Graph API                                                                               |
| As a user, I want to set a message in a chat thread as unread                                                    | P3       | Currently not supported by ACS UI Library                                                                  |
| As a user, I want to pin a message to a chat thread                                                              | P3       | Currently not supported by ACS UI Library                                                                  |


## User Scenarios (out of scope)

Based on feedback and usage, these scenarios could be reprioritized later but are currently considered out-of-scope.

| Scenario                                                                         | Comments |
|----------------------------------------------------------------------------------|----------|
| As a user, I want to see a list of available chats with the last message preview |          |
| As a user, I want to create a new 1:1 chat                                       |          |
| As a user, I want to create a new group chat                                     |          |
| As a user, I want to set a chat thread as unread                                 |          |
| As a user, I want to pin a chat thread to my pinned list                         |          |
| As a user, I want to mute or unmute a chat thread                                |          |

## Proposed Solution

### Examples
```tsx
<Conversation 
    threadId="19:48d31887-5fad-4d73-a9f5-3c356e68a038_b77f6f87-ef1a-42f2-8c71-b5e297e63df3@unq.gbl.spaces">
</Conversation>
```

```tsx
<Conversation 
    threadId="19:48d31887-5fad-4d73-a9f5-3c356e68a038_b77f6f87-ef1a-42f2-8c71-b5e297e63df3@unq.gbl.spaces"
    readOnly={true}>
</Conversation>
```

```tsx
<Conversation 
    conversationId="19:48d31887-5fad-4d73-a9f5-3c356e68a038_b77f6f87-ef1a-42f2-8c71-b5e297e63df3@unq.gbl.spaces"
    theme={fluentTheme}>
</Conversation>
```

##

## Properties and Attributes

| Attribute                 | Property                | Description                                                                                                                                                                                                                           | Type                       |
|---------------------------|-------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------|
| placeholder               | placeholder             | Specify placeholder to use in the combobox. No default value is provided.                                                                                                                                                             | string                     |
| title-property            | titleProperty           | Specify the entity property to use as the title field in the picker. By default, we should use the following : `displayName` / `title` / `name` / `subject` / `id`. Property should exist in the target entity and is case sensitive. | string                     |
| show-max                  | showMax                 | Specify the number of results to show for the resource. Not in use if `max-pages` is used.                                                                                                                                            | Number                     |
| resource                  | resource                | The resource to get from Microsoft Graph (for example, `/me`).                                                                                                                                                                        | String                     |
| scopes                    | scopes                  | Optional array of strings if using the property or a comma delimited scope if using the attribute. The component will use these scopes (with a supported provider) to ensure that the user has consented to the right permission.     | String or array of strings |
| version                   | version                 | Optional API version to use when making the GET request. Default is `v1.0`.                                                                                                                                                           | String                     |
| max-pages                 | maxPages                | Optional number of pages (for resources that support paging). Default is 3. Setting this value to 0 will get all pages.                                                                                                               | Number                     |
| cache-enabled             | cacheEnabled            | Optional Boolean. When set, it indicates that the response from the resource will be cached. Overriden if `refresh()` is called. Default is `false`.                                                                                  | Boolean                    |
| cache-invalidation-period | cacheInvalidationPeriod | Optional number of milliseconds. When set in combination with `cacheEnabled`, the delay before the cache reaches its invalidation period will be modified by this value. Default is `0` and will use the default invalidation period. | Number                     |


## Methods

| Method                  | Description                                                                                                                                 |
|-------------------------|---------------------------------------------------------------------------------------------------------------------------------------------|
| refresh(force?:boolean) | Call the method to refresh the data. By default, the UI will only update if the data changes. Pass `true` to force the component to update. |

## Events

| Event            | When it is fired                                |
|------------------|-------------------------------------------------|
| selectionChanged | Fired when the user makes a change in selection |

## Templates

| Data type     | Data Context                          | Description                                                                                  |
|---------------|---------------------------------------|----------------------------------------------------------------------------------------------|
| default       | null: no data                         | The template used to override the rendering of the entire component.                         |
| loading       | null: no data                         | The template used to render the state of the picker while the request to Graph is being made |
| error         | null: no data                         | The template used if search returns no results.                                              |
| rendered-item | renderedItem: the item being rendered | The template used to render the item inside the dropdown.                                    |

## APIs and Permissions

Permissions required by this component when using the `resource` property depend on the data that you want to retrieve with it from Microsoft Graph. For more information about permissions, see the Microsoft Graph [permissions reference](https://learn.microsoft.com/graph/permissions-reference).