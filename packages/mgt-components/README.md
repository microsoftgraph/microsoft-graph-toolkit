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

      // import the registration functions
      import {
        registerMgtLoginComponent,
        registerMgtAgendaComponent
      } from '@microsoft/mgt-components';

      // register the components
      registerMgtLoginComponent();
      registerMgtAgendaComponent();

      // initialize the auth provider globally
      Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});
    </script>

    <mgt-login></mgt-login>
    <mgt-person person-query="Bill Gates" person-card="hover"></mgt-person>
    <mgt-agenda group-by-day></mgt-agenda>
    ```

> **Important:** Components must now be explicitly registered to function, this is a breaking change in v4.0.0

## Tree shaking/Live code inclusion

By requiring explicit component registration bundlers can now correctly tree shake the MGT libraries. If you use the convenience `registerMgtComponents()` function then all components will be included in your bundle. For production applications you should explicitly register each component that will be used in your application using the appropriate function. For each component there is a registration function `registerMgt{Name}Component()`, e.g. `registerMgtLoginComponent()`.

In cases where a component has a dependency on other components these are all registered in the registration function. For example, the mgt-login component uses an mgt-person internally, so `registerMgtLoginComponent()` calls `registerMgtPersonComponent()`.`

### Why use a explicit function call for component registration?

This removes the registration of a component from being a side effect of importing the component, this allows for imperative disambiguation of web components without the need to rely on dynamic imports to ensure that disambiguation is configured before the import of the `@microsoft/mgt-components` library happens. It also provides for greater developer control.

## <a id="disambiguation">Disambiguation</a>

MGT is built using [web components](https://developer.mozilla.org/en-US/docs/Web/Web_Components). Web components use their tag name as a unique key when registering within a browser. Any attempt to register a component using a previously registered tag name results in an error being thrown when calling `CustomElementRegistry.define()`. In scenarios where multiple custom applications can be loaded into a single page this created issues for MGT, most notably in developing solutions using SharePoint Framework.

To allow developers to build web parts with MGT and load them on pages along with other applications using MGT, we've added a disambiguation feature to MGT. Using this feature, developers can specify a unique string to add to the tag name of all MGT web components in their application.


### Usage in standard HTML and JavaScript

The earlier example can be updated to use the disambiguation feature as follows:

```html
<script type="module">
  import { Providers, customElementHelper } from '@microsoft/mgt-element';
  import { Msal2Provider } from '@microsoft/mgt-msal2-provider';
  import {
    registerMgtLoginComponent,
    registerMgtAgendaComponent
  } from '@microsoft/mgt-components';

  // configure disambiguation
  customElementHelper.withDisambiguation('contoso');

  // register the components
  registerMgtLoginComponent();
  registerMgtAgendaComponent();

  // initialize the auth provider globally
  Providers.globalProvider = new Msal2Provider({clientId: 'clientId'});
</script>

<mgt-contoso-login></mgt-contoso-login>
<mgt-contoso-person person-query="Bill Gates" person-card="hover"></mgt-contoso-person>
<mgt-contoso-agenda group-by-day></mgt-contoso-agenda>
```

> Note: `withDisambiguation('foo')` must be called before registering the the desired components.

When developing SharePoint Framework web parts the pattern for using disambiguation is based on whether or not the MGT React wrapper library is being used. If you are using React then the helper utility in the [`mgt-spfx-utils`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx-utils) package can be used. SharePoint Framework web part example usages are provided below.

### Usage in a SharePoint web part with no framework

This example is sourced from the [No Framework Web Part Sample](https://github.com/pnp/mgt-samples/blob/main/samples/app/sp-mgt/src/webparts/helloWorld/HelloWorldWebPart.ts).

```ts
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Providers, customElementHelper } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { registerMgtLoginComponent } from '@microsoft/mgt-components';

export default class MgtWebPart extends BaseClientSideWebPart<Record<string, unknown>> {

  protected onInit(): Promise<void> {
    customElementHelper.withDisambiguation('foo');

    // register the component
    registerMgtLoginComponent();

    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    return super.onInit();
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

When using Microsoft Graph Toolkit components via the `@microsoft/mgt-react` wrapper the underlying custom element registration is handled when generating the wrapping React function component, meaning that the register calls are made when each component is imported. As such the pattern for using disambiguation is unchanged with the introduction of tree shaking compatibility.

The `lazyLoadComponent` helper function from [`mgt-spfx-utils`](https://github.com/microsoftgraph/microsoft-graph-toolkit/tree/main/packages/mgt-spfx-utils) leverages `React.lazy` and `React.Suspense` to asynchronously load the components which have a direct dependency on `@microsoft/mgt-react` from the top level web part component.

A complete example is available in the [React SharePoint Web Part Sample](https://github.com/microsoftgraph/microsoft-graph-toolkit/blob/main/samples/sp-webpart/src/webparts/mgtDemo/MgtDemoWebPart.ts).

```ts
// [...] trimmed for brevity
import { Providers, customElementHelper } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
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

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)