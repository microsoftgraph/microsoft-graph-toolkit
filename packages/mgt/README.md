# Microsoft Graph Toolkit

<a href="https://www.npmjs.com/package/@microsoft/mgt"><img src="https://img.shields.io/npm/v/@microsoft/mgt.svg"></a> <a href="https://github.com/microsoftgraph/msgraph-sdk-javascript"><img src="https://cdn.jsdelivr.net/gh/storybookjs/brand@master/badge/badge-storybook.svg"></a>

The Microsoft Graph Toolkit is a collection of reusable, framework-agnostic components and authentication providers for accessing and working with Microsoft Graph. The components are fully functional right of out of the box, with built in providers that authenticate with and fetch data from Microsoft Graph.

The `@microsoft/mgt` package brings all mgt packages together (with the exception of `@microsoft/mgt-react`) and bundles them in this one convenient package.

[View the full documentation](https://learn.microsoft.com/graph/toolkit/overview)

You can now explore components and samples with the [playground](https://mgt.dev).
## Components

The Microsoft Graph Toolkit includes a collection of web components for the most commonly built experiences powered by Microsoft Graph APIs.

The components are also available as [React components](https://learn.microsoft.com/graph/toolkit/get-started/mgt-react).

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

## Providers
[Providers](https://learn.microsoft.com/graph/toolkit/providers) enable authentication and provide the implementation for acquiring access tokens on various platforms and expose a Microsoft Graph Client for calling the Microsoft Graph APIs. The components work best when used with a provider, but the providers can be used on their own.

* [Msal2Provider](https://learn.microsoft.com/graph/toolkit/providers/msal2)
* [SharePointProvider](https://learn.microsoft.com/graph/toolkit/providers/sharepoint)
* [TeamsFxProvider](https://learn.microsoft.com/graph/toolkit/providers/teamsfx)
* [ProxyProvider](https://learn.microsoft.com/graph/toolkit/providers/proxy)
* [SimpleProvider](https://learn.microsoft.com/graph/toolkit/providers/custom)
* [ElectronProvider](https://learn.microsoft.com/graph/toolkit/providers/electron)


## Getting Started

[Watch the Getting Started Video](https://www.youtube.com/watch?v=oZCGb2MMxa0)

You can use the components by installing the npm package or importing them from a CDN (unpkg).

### Use via NPM:

The benefits of using MGT through NPM is that you have full control of the bundling process and you can bundle only the code you need for your site. First, add the npm package:

```bash
npm install @microsoft/mgt-components
npm install @microsoft/mgt-msal2-provider
```

Now you can reference all components and providers at the page you are using:

```html
<script type="module">
  import { Providers } from 'node_modules/@microsoft/mgt-element/dist/es6/index.js';
  import { Msal2Provider } from 'node_modules/@microsoft/mgt-msal2-provider/dist/es6/index.js';
  import { registerMgtLoginComponent, registerMgtAgendaComponent } from 'node_modules/@microsoft/mgt-components/dist/es6/index.js';
  
  Providers.globalProvider = new Msal2Provider({clientId: '[CLIENT-ID]'});
  
  registerMgtLoginComponent();
  registerMgtAgendaComponent();
</script>

<mgt-login></mgt-login>
<mgt-agenda></mgt-agenda>
```

### Use via CDN:

The following script tag downloads the code from the CDN, configures an MSAL2 provider, and makes all the components available for use in the web page.

```html
<script type="module">
  import { registerMgtComponents, Providers, Msal2Provider } from 'https://unpkg.com/@microsoft/mgt@4';
  Providers.globalProvider = new Msal2Provider({clientId: '[CLIENT-ID]'});
  registerMgtComponents();
</script>
<mgt-login></mgt-login>
<mgt-agenda></mgt-agenda>
```

> NOTE: This link will load the highest available version of @microsoft/mgt in the range `>= 4.0.0 < 5.0.0`, omitting the `@4` fragment from the url results in loading the latest version. This could result in loading a new major version and breaking the application.

> NOTE: MSAL requires the page to be hosted in a web server for the authentication redirects. If you are just getting started and want to play around, the quickest way is to use something like [live server](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer) in vscode.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)