---
title: "Chat component in Microsoft Graph Toolkit"
description: "The Chat component enables the user to have full 1:1 or group conversations. It works with any conversation happening in Microsoft Teams."
ms.localizationpriority: medium
author: sebastienlevert
---

# Chat component in Microsoft Graph Toolkit

[!IMPORTANT] This component is in Preview and is subject to change. The use of these components in production applications is not supported.

[!NOTE] This component is only available as a React component but this is subject to change.

The Chat component enables the user to have full 1:1 or group conversations. It works with any conversation happening in Microsoft Teams.

## Example

The following example displays a signed-in user's conversation using the `mgt-chat` component. You can use the code editor to see how [properties](#properties) change the behavior of the component.

<iframe src="https://mgt.dev/iframe.html?id=components-mgt-chat--chat&source=docs" height="500"></iframe> -->

[Open this example in mgt.dev](https://mgt.dev/?path=/story/components-mgt-chat--chat&source=docs)

## Properties

| Attribute                         | Property         | Description                                                                                            |
| --------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------ |
| chat-id                         | chatId         | A string ID to set the 1:1 or group conversation to render. Required. |

The following example shows a conversation with ID _19:25fdc88d202440b78e9229773cbb1713@thread.v2_.

```tsx
<Chat chatId={'19:25fdc88d202440b78e9229773cbb1713@thread.v2'} />
```

## CSS custom properties

The `mgt-chat` component does not define CSS custom properties.

## Events

The `mgt-chat` component doesn't offer any events.

## Templates

The `mgt-chat` component doesn't offer any template to override.

## Microsoft Graph permissions

This control uses the following Microsoft Graph APIs and permissions.

| Configuration | Permission | API |
| - | - | - |
| `chatId` is set | Chat.ReadBasic, Chat.Read, ChatMessage.Read, Chat.ReadWrite, ChatMember.ReadWrite | [/chats/{id}/messages](https://learn.microsoft.com/graph/api/chat-list-messages?view=graph-rest-1.0&tabs=http), [/chats/{id}/messages](https://learn.microsoft.com/graph/api/chat-post-messages?view=graph-rest-1.0&tabs=http), [/chats/{id}/messages/{messageId}](https://learn.microsoft.com/graph/api/chatmessage-update?view=graph-rest-1.0&tabs=http), [/me/chats/{id}/messages/{messageId}/softDelete](https://learn.microsoft.com/graph/api/chatmessage-softdelete?view=graph-rest-1.0&tabs=http), [/chats/{id}/members/{membershipId}](https://learn.microsoft.com/graph/api/chat-delete-members?view=graph-rest-1.0&tabs=http), [/chats/{id}/members](https://learn.microsoft.com/graph/api/chat-post-members?view=graph-rest-1.0&tabs=http), [/chats/{id}/messages/{messageId}/hostedContents/{hostedContentId}](https://learn.microsoft.com/graph/api/chatmessagehostedcontent-get?view=graph-rest-1.0&tabs=http), [/chats/{id}](https://learn.microsoft.com/graph/api/chat-patch?view=graph-rest-1.0&tabs=http) |

## Authentication

The tasks component uses the global authentication provider described in the [authentication documentation](../providers/providers.md).

## Cache

The `mgt-tasks` component doesn't cache any data.

## Localization

The `mgt-tasks` component doesn't expose any localization variables.
