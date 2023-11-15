# mgt-chat

## Overview

The chat component enables the user to have 1:1 or group conversations. This component doesn't support channel conversations. The component allows for rendering conversations and authoring new messages. All data is stored in Microsoft Teams.

## User scenarios

### Priority 0

- [Supporting 1:1 conversations](#000)
- [Supporting group conversations](#001)
- [Receiving real-time messages from Teams](#002)
- [Ensuring all system messages are rendered in the component](#003)
- [Ensuring DLP policies are respected in the component](#004)
- [Managing the roster of a group conversation](#005)
- [Managing the title of the conversation (when a group conversation)](#006)
- [Navigating to Microsoft Teams from the component](#007)
- [Rendering the most used content types in the component (rich text, emojis, images, GIFs, mentions, links, adaptive cards)](#008)
- [Rendering an unsupported card in the component when the component doesn't know about the content type](#009)
- [Displaying avatars, presence and cards from the users sending messages](#010)
- [Enabling multiple chat threads in the same page](#011)
-

### Priority 1

- Supporting custom themes to match the host application
- Displaying large status messages to users when something goes wrong
- Allowing users to mention members of the conversation
- Improve performance of the Graph calls by leveraging premium APIs
- Supporting developers in other languages than React

### Priority 2

- Rendering read receipts
- Rendering of reactions authored in Teams
- Reacting to messages from the component
- Identifying that a user is currently typing
- Formatting fich content in the authoring experience
- Rendering of attached files added in Teams
- Attaching images in the authoring experience
- Attaching files in the authoring experience
- Rendering and interacting with channel conversations

## Requirements

### Supporting 1:1 conversations <a id="000" />

### Supporting group conversations <a id="001" />

### Receiving real-time messages from Teams <a id="002" />

### Ensuring all system messages are rendered in the component  <a id="003" />

### Ensuring DLP policies are respected in the component  <a id="004" />

### Managing the roster of a group conversation  <a id="005" />

### Managing the title of the conversation (when a group conversation)  <a id="006" />

### Navigating to Microsoft Teams from the component  <a id="007" />

### Rendering the most used content types in the component (rich text, emojis, images, GIFs, mentions, links, adaptive cards)  <a id="008" />

### Rendering an unsupported card in the component when the component doesn't know about the content type  <a id="009" />

### Displaying avatars, presence and cards from the users sending messages  <a id="010" />

### Enabling multiple chat threads in the same page  <a id="011" />

## Designs

All the designs for this component can be found [here](https://www.figma.com/file/gP4q8VQ2so2ftzKz4GCuVh/Microsoft-Graph-Toolkit-(WIP)?node-id=10197%3A127838&mode=dev).

## Advanced Examples

```tsx
<Chat chatId="19:2894b8353ac24fbbbc13391aab6fc952@thread.v2" />
```

```tsx
<Chat usePremiumApis={true} chatId="19:2894b8353ac24fbbbc13391aab6fc952@thread.v2" />
```

## Properties

| Attribute                         | Property         | Description                                                                                            |
| --------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------ |
| chat-id                           | chatId           | A string ID to set the 1:1 or group [conversation](/graph/api/resources/chat) to render. Required.     |
| use-premium-apis                           | usePremiumApis           | A boolean value representing if this component is enabled to call the Premium APIs. Optional.     |

## CSS custom properties

The `mgt-chat` component doesn't define CSS custom properties.

## Events

The `mgt-chat` component doesn't offer any events.

## Templates

The `mgt-chat` component doesn't offer templates to override.

## Microsoft Graph permissions

This control uses the following Microsoft Graph APIs and permissions.

| Configuration | Permission | API |
| - | - | - |
| `chatId` is set | Chat.ReadBasic, Chat.Read, ChatMessage.Read, Chat.ReadWrite, ChatMember.ReadWrite | [/chats/{id}/messages](/graph/api/chat-list-messages), [/chats/{id}/messages](/graph/api/chat-post-messages), [/chats/{id}/messages/{messageId}](/graph/api/chatmessage-update), [/me/chats/{id}/messages/{messageId}/softDelete](/graph/api/chatmessage-softdelete), [/chats/{id}/members/{membershipId}](/graph/api/chat-delete-members), [/chats/{id}/members](/graph/api/chat-post-members), [/chats/{id}/messages/{messageId}/hostedContents/{hostedContentId}](/graph/api/chatmessagehostedcontent-get), [/chats/{id}](/graph/api/chat-patch) |

## Authentication

The `mgt-chat` component uses the global authentication provider described in the [authentication documentation](../providers/providers.md).

## Cache

The `mgt-chat` component caches chat messages and related metadata.

## Localization

The `mgt-chat` component doesn't expose any localization variables.
