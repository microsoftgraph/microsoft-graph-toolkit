# Microsoft Graph Toolkit Web Components

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-components?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-components)

The [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) web components package is a collection of web components powered by Microsoft Graph. The components are functional, work automatically with Microsoft Graph, and work with any web framework and on all modern browsers.

> Note: If you are building with React, you might find `@microsoft/mgt-react` useful as well.

[See docs for full documentation](https://aka.ms/mgt-docs)

## Components
You can explore components and samples with the [playground](https://mgt.dev) powered by storybook.

The Toolkit currently includes the following components:

* [mgt-agenda](https://learn.microsoft.com/graph/toolkit/components/agenda)
* [mgt-file](https://learn.microsoft.com/graph/toolkit/components/file)
* [mgt-file-list](https://learn.microsoft.com/graph/toolkit/components/file-list)
* [mgt-get](https://learn.microsoft.com/graph/toolkit/components/get)
* [mgt-login](https://learn.microsoft.com/graph/toolkit/components/login)
* [mgt-people](https://learn.microsoft.com/graph/toolkit/components/people)
* [mgt-people-picker](https://learn.microsoft.com/graph/toolkit/components/people-picker)
* [mgt-person](https://learn.microsoft.com/graph/toolkit/components/person)
* [mgt-person-card](https://learn.microsoft.com/graph/toolkit/components/person-card)
* [mgt-picker](https://learn.microsoft.com/en-us/graph/toolkit/components/picker)
* [mgt-search-box](https://learn.microsoft.com/graph/toolkit/components/person-box)
* [mgt-search-results](https://learn.microsoft.com/graph/toolkit/components/search-results)
* [mgt-tasks](https://learn.microsoft.com/graph/toolkit/components/tasks)
* [mgt-taxonomy-picker](https://learn.microsoft.com/graph/toolkit/components/taxonomy-picker)
* [mgt-teams-channel-picker](https://learn.microsoft.com/graph/toolkit/components/teams-channel-picker)
* [mgt-theme-toggle](https://learn.microsoft.com/graph/toolkit/components/theme-toggle)
* [mgt-todo](https://learn.microsoft.com/graph/toolkit/components/todo)

The components work best when used with a [provider](https://learn.microsoft.com/graph/toolkit/providers). The provider handles authentication and the requests to the Microsoft Graph APIs used by the components.

## Get started

The components can be used on their own, but they are at their best when they are paired with an authentication provider. This example illustrates how to use the components alongside a provider (MsalProvider in this case).

1. Install the packages

    ```bash
    npm install @microsoft/mgt-element @microsoft/mgt-components @microsoft/mgt-msal2-provider
    ```

1. Use components in your code

    ```html
    <script type="module">
      import {Providers} from '@microsoft/mgt-element';
      import {Msal2Provider} from '@microsoft/mgt-msal2-provider';

      // import the components
      import '@microsoft/mgt-components';

      // initialize the auth provider globally
      Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});
    </script>

    <mgt-login></mgt-login>
    <mgt-person person-query="Bill Gates" person-card="hover"></mgt-person>
    <mgt-agenda group-by-day></mgt-agenda>
    ```


## <a id="disambiguation">Disambiguation</a>

MGT is built using [web components](https://developer.mozilla.org/en-US/docs/Web/Web_Components). Web components use their tag name as a unique key when registering within a browser. Any attempt to register a component using a previously registered tag name results in an error being thrown when calling `CustomElementRegistry.define()`. In scenarios where multiple custom applications can be loaded into a single page this created issues for MGT, most notably in developing solutions using SharePoint Framework.

To mitigate this challenge we built the [`mgt-spfx`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx) package. Using `mgt-spfx` developers can centralize the registration of MGT web components across all SPFx solutions deployed on the tenant. By reusing MGT components from a central location web parts from different solutions can be loaded into a single page without throwing errors. When using `mgt-spfx` all MGT based web parts in a SharePoint tenant use the same version of MGT.

To allow developers to build web parts using the latest version of MGT and load them on pages along with web parts that use v2.x of MGT, we've added a new disambiguation feature to MGT. Using this feature developers can specify a unique string to add to the tag name of all MGT web components in their application.

### Usage in standard HTML and JavaScript

The earlier example can be updated to use the disambiguation feature as follows:

```html
<script type="module">
  import { Providers, customElementHelper } from '@microsoft/mgt-element';
  import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
  // configure disambiguation
  customElementHelper.withDisambiguation('contoso');

  // initialize the auth provider globally
  Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});

  // import the components using dynamic import to avoid hoisting
  import('@microsoft/mgt-components');
</script>

<mgt-contoso-login></mgt-contoso-login>
<mgt-contoso-person person-query="Bill Gates" person-card="hover"></mgt-contoso-person>
<mgt-contoso-agenda group-by-day></mgt-contoso-agenda>
```

> Note: the `import` of `mgt-components` must use a dynamic import to ensure that the disambiguation is applied before the components are imported.

When developing SharePoint Framework web parts the pattern for using disambiguation is based on whether or not the MGT React wrapper library is being used. If you are using React then the helper utility in the [`mgt-spfx-utils`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx-utils) package should be used. SharePoint Framework web part example usages are provided below.

### Usage in a SharePoint web part with no framework

A dynamic import of the `@microsoft/mgt-components` library is used after configuring the disambiguation.

This example is sourced from the [No Framework Web Part Sample](https://github.com/microsoftgraph/microsoft-graph-toolkit/blob/main/samples/sp-mgt/src/webparts/helloWorld/HelloWorldWebPart.ts).

```ts
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';

export default class MgtWebPart extends BaseClientSideWebPart<Record<string, unknown>> {

  protected onInit(): Promise<void> {
    customElementHelper.withDisambiguation('foo');
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    return import('@microsoft/mgt-components').then(() => super.onInit());
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <mgt-foo-login></mgt-foo-login>
    </section>`;
  }
}
```

### Usage in a SharePoint web part using React

The `lazyLoadComponent` helper function from [`mgt-spfx-utils`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx-utils) leverages `React.lazy` and `React.Suspense` to asynchronously load the components which have a direct dependency on `@microsoft/mgt-react` from the top level web part component.

A complete example is available in the [React SharePoint Web Part Sample](https://github.com/microsoftgraph/microsoft-graph-toolkit/blob/main/samples/sp-webpart/src/webparts/mgtDemo/MgtDemoWebPart.ts).

```ts
// [...] trimmed for brevity
import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';
import { lazyLoadComponent } from '@microsoft/mgt-spfx-utils';

// Async import of a component that uses the @microsoft/mgt-react Components
const MgtDemo = React.lazy(() => import('./components/MgtDemo'));

export interface IMgtDemoWebPartProps {
  description: string;
}
// set the disambiguation before initializing any webpart
customElementHelper.withDisambiguation('bar');

export default class MgtDemoWebPart extends BaseClientSideWebPart<IMgtDemoWebPartProps> {
  // set the global provider
  protected async onInit() {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
  }

  public render(): void {
    const element = lazyLoadComponent(MgtDemo, { description: this.properties.description });

    ReactDom.render(element, this.domElement);
  }

  // [...] trimmed for brevity
}
```

### Dynamic imports aka Lazy Loading

Using dynamic imports you can load dependencies asynchronously. This pattern allows you to load dependencies only when needed. For example, you may want to load a component only when a user clicks a button. This is a great way to reduce the initial load time of your application. In the context of disambiguation, you need to use this technique because components register themselves in the browser when they are imported.

**Important:** If you import the components before you have applied the disambiguation, the disambiguation will not be applied and using the disambiguated tag name will not work.

When using an `import` statement the import statement is hoisted and executed before any other code in the code block. To use dynamic imports you must use the `import()` function. The `import()` function returns a promise that resolves to the module. You can also use the `then` method to execute code after the module is loaded and the `catch` method to handle any errors if necessary.

#### Example using dynamic imports

```typescript
// static import via a statement
import { Providers, customElementHelper } from '@microsoft/mgt-element';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';

customElementHelper.withDisambiguation('contoso');
Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});

// dynamic import via a function
import('@microsoft/mgt-components').then(() => {
  // code to execute after the module is loaded
  document.body.innerHTML = '<mgt-contoso-login></mgt-contoso-login>';
}).catch((e) => {
  // handle any errors
});
```

#### Example using static imports

```typescript
// static import via a statement
import { Provider } from '@microsoft/mgt-element';
import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
import '@microsoft/mgt-components';

Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});

document.body.innerHTML = '<mgt-login></mgt-login>';
```

> Note: it is not possible to use disambiguation with static imports.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)