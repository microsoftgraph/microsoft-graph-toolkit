---
title: "New Chat component in Microsoft Graph Toolkit"
description: "The New Chat component allows user to create new 1:1 or group conversations in Microsoft Teams."
ms.localizationpriority: medium
author: sebastienlevert
---

# New Chat component in Microsoft Graph Toolkit

[!IMPORTANT] This component is in Preview and is subject to change. The use of these components in production applications is not supported.

[!NOTE] This component is only available as a React component but this is subject to change.

The New Chat component allows user to create new 1:1 or group conversations in Microsoft Teams.

## Example

The following example displays a signed-in user's conversation using the `mgt-new-chat` component. You can use the code editor to see how [properties](#properties) change the behavior of the component.

<iframe src="https://mgt.dev/iframe.html?id=components-mgt-new-chat--new-chat&source=docs" height="500"></iframe> -->

[Open this example in mgt.dev](https://mgt.dev/?path=/story/components-mgt-new-chat--new-chat&source=docs)

## Properties

| Attribute | Property | Description |
| - | - | - |
| mode | mode | Set to `oneOnOne`, `group` or `auto`. Default is `auto`. |

## CSS custom properties

The `mgt-new-chat` component does not define CSS custom properties.

## Events

The following events are fired from the component.

Event | When is it emitted | Custom data | Cancelable | Bubbles | Works with custom template
------|-------------------|--------------|:-----------:|:---------:|:---------------------------:|
`onChatCreated` | Fired when a new chat thread is created. | The `chat` object that was created as a Microsoft Graph [chat](https://learn.microsoft.com/graph/api/resources/chat?view=graph-rest-1.0#json-representation). | No | No | No |
`onCancelClicked` | Fired when user cancels the chat thread creation. | None | No | No | No |

For more information about handling events, see [events](../customize-components/events.md).

## Templates

The `mgt-chat` component doesn't offer any template to override.

## Microsoft Graph permissions

This control uses the following Microsoft Graph APIs and permissions.

| Configuration | Permission | API |
| - | - | - |
| Default | Chat.Create, ChatMessage.Send | [/chats](https://learn.microsoft.com/graph/api/chat-post?view=graph-rest-1.0&tabs=http), [/chats/{id}/messages](https://learn.microsoft.com/graph/api/chat-post-messages?view=graph-rest-1.0&tabs=http) |

## Authentication

The tasks component uses the global authentication provider described in the [authentication documentation](../providers/providers.md).

## Cache

The `mgt-tasks` component doesn't cache any data.

## Localization

The `mgt-tasks` component doesn't expose any localization variables.
