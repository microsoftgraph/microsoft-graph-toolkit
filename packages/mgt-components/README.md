# Microsoft Graph Toolkit Web Components

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-components?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-components)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) web components package is a collection of web components powered by Microsoft Graph. The components are functional, work automatically with Microsoft Graph, and work with any web framework and on all modern browsers.

> Note: If you are building with React, you might find `@microsoft/mgt-react` useful as well.

[See docs for full documentation](https://aka.ms/mgt-docs)

## Components
You can explore components and samples with the [playground](https://mgt.dev) powered by storybook.

The Toolkit currently includes the following components:

* [mgt-login](https://docs.microsoft.com/graph/toolkit/components/login)
* [mgt-person](https://docs.microsoft.com/graph/toolkit/components/person)
* [mgt-person-card](https://docs.microsoft.com/graph/toolkit/components/person-card)
* [mgt-people](https://docs.microsoft.com/graph/toolkit/components/people)
* [mgt-people-picker](https://docs.microsoft.com/graph/toolkit/components/people-picker)
* [mgt-agenda](https://docs.microsoft.com/graph/toolkit/components/agenda)
* [mgt-tasks](https://docs.microsoft.com/graph/toolkit/components/tasks)
* [mgt-todo](https://docs.microsoft.com/graph/toolkit/components/todo)
* [mgt-get](https://docs.microsoft.com/graph/toolkit/components/get)
* [mgt-teams-channel-picker](https://docs.microsoft.com/en-us/graph/toolkit/components/teams-channel-picker)

The components work best when used with a [provider](https://docs.microsoft.com/graph/toolkit/providers). The provider handles authentication and the requests to the Microsoft Graph APIs used by the components.

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
      import {Msal2Provider} from '@microsoft/mgt-msal2-provider';

      // import the components
      import '@microsoft/mgt-components';

      // initialize the auth provider globally
      Providers.globalProvider = new MsalProvider({clientId: 'clientId'});
    </script>

    <mgt-login></mgt-login>
    <mgt-person person-query="Bill Gates" person-card="hover"></mgt-person>
    <mgt-agenda group-by-day></mgt-agenda>
    ```


## <a id="disambiguation">Disambiguation</a>

MGT is built using [web components](https://developer.mozilla.org/en-US/docs/Web/Web_Components). Web components use their tag name as a unique key when registering within a browser. Any attempt to register a component using a previously registered tag name results in an error being thrown when calling `CustomElementRegistry.define()`. In scenarios where multiple custom applications can be loaded into a single page this created issues for MGT, most notably in developing custom SharePoint web parts.

To mitigate this challenge we built the [`mgt-spfx`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx) package. Using `mgt-spfx` developers could centralize the registration of MGT web components across all SPFx solutions deployed on the tenant. By reusing MGT components from a central location web parts from different solutions could be loaded into a single page without throwing errors. When using `mgt-spfx` all MGT based web parts in a SharePoint tenant used the same version of MGT.

To allow developers to build web parts using the latest version of MGT and load them on pages along web parts that use v2.x of MGT, we've added a new disambiguation feature to MGT. Using this feature developers can specify a unique string to add to the tag name of all MGT web components in their application.

The earlier example can be updated to use the disambiguation feature as follows:

```html
<script type="module">
  import { Providers, customElementHelper } from '@microsoft/mgt-element';
  customElementHelper.withDisambiguation('contoso');
  import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
  Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});

</script>

<mgt-contoso-login></mgt-contoso-login>
<mgt-contoso-person person-query="Bill Gates" person-card="hover"></mgt-contoso-person>
<mgt-contoso-agenda group-by-day></mgt-contoso-agenda>
```
> Note: the `import` of `mgt-components` must use a dynamic import to ensure that the disambiguation is applied before the components are imported. Helper utilities have been provided in the  [`mgt-spfx-utils`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx-utils) package to make this easier in the context of SharePoint web parts.

### Dynamic imports aka Lazy Loading

Using dynamic imports you can load dependencies asynchronously. This pattern allows you to load dependencies only when needed. For example, you may want to load a component only when a user clicks a button. This is a great way to reduce the initial load time of your application. In the context of disambiguation, you need to use this technique becasuse components register themselves in the browser when they are imported.

When using an `import` statement the import statement is hoisted and executed before any other code in the code block. To use dynamic imports you must use the `import()` function.

**Important:** If you import the components before you have applied the disambiguation, the disambiguation will not be applied and using the disambiguated tag name will not work.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)