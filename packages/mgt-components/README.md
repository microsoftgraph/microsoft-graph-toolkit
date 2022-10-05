# Microsoft Graph Toolkit Web Components

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-components?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-components)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) web components package is a collection of web components powered by Microsoft Graph. The components are functional, work automatically with Microsoft Graph, and work with any web framework and on all modern browsers.

> Note: If you are building with React, you might find `@microsoft/mgt-react` useful as well.

[See docs for full documentation](https://aka.ms/mgt-docs)

## Components
You can explore components and samples with the [playground](https://mgt.dev) powered by storybook.

The Toolkit currently includes the following components:

* [mgt-login](https://learn.microsoft.com/graph/toolkit/components/login)
* [mgt-person](https://learn.microsoft.com/graph/toolkit/components/person)
* [mgt-person-card](https://learn.microsoft.com/graph/toolkit/components/person-card)
* [mgt-people](https://learn.microsoft.com/graph/toolkit/components/people)
* [mgt-people-picker](https://learn.microsoft.com/graph/toolkit/components/people-picker)
* [mgt-agenda](https://learn.microsoft.com/graph/toolkit/components/agenda)
* [mgt-tasks](https://learn.microsoft.com/graph/toolkit/components/tasks)
* [mgt-todo](https://learn.microsoft.com/graph/toolkit/components/todo)
* [mgt-get](https://learn.microsoft.com/graph/toolkit/components/get)
* [mgt-teams-channel-picker](https://learn.microsoft.com/graph/toolkit/components/teams-channel-picker)

The components work best when used with a [provider](https://learn.microsoft.com/graph/toolkit/providers). The provider handles authentication and the requests to the Microsoft Graph APIs used by the components.

## Get started

The components can be used on their own, but they are at their best when they are paired with an authentication provider. This example illustrates how to use the components alongside a provider (MsalProvider in this case).

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-components @microsoft/mgt-msal-provider
    ```

1. Use components in your code

    ```html
    <script type="module">
      import {Providers} from '@microsoft/mgt-element';
      import {MsalProvider} from '@microsoft/mgt-msal-provider';

      // import the components
      import '@microsoft/mgt-components';

      // initialize the auth provider globally
      Providers.globalProvider = new MsalProvider({clientId: 'clientId'});
    </script>

    <mgt-login></mgt-login>
    <mgt-person person-query="Bill Gates" person-card="hover"></mgt-person>
    <mgt-agenda group-by-day></mgt-agenda>
    ```

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)