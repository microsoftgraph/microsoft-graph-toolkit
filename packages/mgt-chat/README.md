# Microsoft Graph Toolkit Chat Components

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-chat/next.mgt-chat?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-chat)

> Note: The Microsoft Graph Toolkit Chat Components are in Private Preview and are subject to change. The use of these components in production applications is not supported.

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) library is a collection of authentication providers and UI components powered by Microsoft Graph.

The `@microsoft/mgt-chat` package contains all chat components that allow to build chat-based applications with React.

[See docs for full documentation](https://aka.ms/mgt/docs)

[See docs for chat components](https://aka.ms/mgt/docs/chat/components)

## Getting started with the Graph Toolkit chat components

Use this guide to get started with using the Microsoft Graph Toolkit Chat Components. Some of the steps highlights in this guide are temporary and will evolve with the future preview versions of the components.

### Install the Graph Notification Broker

To start, the Chat Components require a server side dependency to help with receiving live notifications from the service. This helps with getting new messages, roster changes, etc. This component is the [Graph Notification Broker](https://github.com/microsoft/GraphNotificationBroker#graph-notification-broker).

To install the Graph Notification Broker, refer to this [deployment guide](https://github.com/microsoft/GraphNotificationBroker/blob/main/docs/deployment-guide.md#Deployment-Guide).

From this deployment, you will need to keep track of 3 created artifacts:

* The URL of the provisioned Azure Function
* The ID of the Backend App Registration
* The ID of the Frontend App Registration

### Install the necessary dependencies

In your React app where you want to add the Microsoft Graph Toolkit Chat Components, you will need to install our dependencies:

```bash
npm install @microsoft/mgt-chat@next.mgt-chat
npm install @microsoft/mgt-react
npm install @microsoft/mgt-msal2-provider
```

### Configure the broker settings and authentication

At the root of your component hierarchy (usually in `index.tsx`), configure the broker settings with the previously noted values. You will also configure the authentication to enable seemless integration with both the Graph Notification Broker and with Microsoft Graph.

```tsx
import { Providers } from '@microsoft/mgt-react';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import { allChatScopes, brokerSettings } from '@microsoft/mgt-chat';

brokerSettings.functionHost = '%URL_AZURE_FUNCTION%';
brokerSettings.appId = '%BACKEND_APP_ID%';

Providers.globalProvider = new Msal2Provider({
  clientId: '%FRONTEND_APP_ID%',
  scopes: allChatScopes
});
```

### Add the components to your code

You have all the necessary settings in place to add the Microsoft Graph Toolkit Chat Components to your application. To add the components, just do the following:

```tsx
import React, { memo, useCallback, useState } from 'react';
import { Login } from '@microsoft/mgt-react';
import { Chat, NewChat } from '@microsoft/mgt-chat';

export const App: React.FunctionComponent = () => {
  const [chatId, setChatId] = useState<string>();

  const [showNewChat, setShowNewChat] = useState<boolean>(false);
  const onChatCreated = useCallback((chat: GraphChat) => {
    setChatId(chat.id);
    setShowNewChat(false);
  }, []);

  return (
    <Login />
    <button onClick={() => setShowNewChat(true)}>New Chat</button>
    {showNewChat && (
        <NewChat
            onChatCreated={onChatCreated}
            onCancelClicked={() => setShowNewChat(false)}
            mode="auto"
        />
    )}

    {chatId && <Chat chatId={chatId} />}
  );
}
```

### Review existing samples

We are providing simple apps that provides all the necessary elements to have a fully working application with our Chat Components.

* [Our basic `react-chat` application](../../samples/react-chat)
* [Our hero `react-contoso` application](../../samples/react-contoso)

### Known issues

* When using the chat components, users need to login twice because it requires to acquire tokens to both Microsoft Graph and the Graph Notification Broker.
* Some messages are duplicating.

## Contribute

We enthusiastically welcome contributions and feedback. Please read our [wiki](https://github.com/microsoftgraph/microsoft-graph-toolkit/wiki) and the [contributing guide](../../CONTRIBUTING.md) before you begin.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

All files in this GitHub repository are subject to the [MIT license](https://github.com/microsoftgraph/microsoft-graph-toolkit/blob/main/LICENSE). This project also references fonts and icons from a CDN, which are subject to a separate [asset license](https://static2.sharepointonline.com/files/fabric/assets/license.txt).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
