<h1 align="center">
  <a href="#"><img height="120" src="https://github.com/microsoftgraph/microsoft-graph-toolkit/raw/main/assets/graff.png" title="Graff the Giraffe"></a>
  <br>
  Microsoft Graph Toolkit
</h1>

<h4 align="center">UI Components and Authentication Providers for <a href="https://graph.microsoft.com">Microsoft Graph</a></h4>

<p align="center">
  <a href="https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit" target="_blank" rel="noreferrer noopener"><img src="https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg"></a>
  <img src="https://github.com/microsoftgraph/microsoft-graph-toolkit/workflows/Build%20CI/badge.svg" /> <a href="https://mgt.dev" target="_blank" rel="noreferrer noopener"><img src="https://cdn.jsdelivr.net/gh/storybookjs/brand@master/badge/badge-storybook.svg"></a> <a href="https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22" target="_blank" rel="noreferrer noopener"><img src="https://img.shields.io/github/issues/microsoftgraph/microsoft-graph-toolkit/good%20first%20issue?color=brightgreen"></a>
</p>

<h3 align="center"><a href="https://aka.ms/mgt/docs" target="_blank" rel="noreferrer noopener">Documentation</a></h3>

<p align="center">
  The Microsoft Graph Toolkit is a collection of reusable, framework-agnostic components and authentication providers for accessing and working with Microsoft Graph. The components are fully functional right of out of the box, with built in providers that authenticate with and fetch data from Microsoft Graph.
</p>

<p align="center">
  <a href="#packages">Packages</a> • <a href="#components">Components</a> • <a href="#providers">Providers</a> • <a href="#getting-started">Getting Started</a> • <a href="#using-our-samples">Using our samples</a> • <a href="#contribute">Contribute</a> • <a href="#feedback-and-requests">Feedback & Requests</a> <br>• <a href="#license">License</a> • <a href="#code-of-conduct">Code of Conduct</a>
</p>

## Packages

| Package                                                                                                    | Latest                                                                                  | Next                                                                                  |
| ---------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------- |
| [`@microsoft/mgt`](https://www.npmjs.com/package/@microsoft/mgt)                                           | <img src="https://img.shields.io/npm/v/@microsoft/mgt/latest.svg">                      | <img src="https://img.shields.io/npm/v/@microsoft/mgt/next.svg">                      |
| [`@microsoft/mgt-element`](https://www.npmjs.com/package/@microsoft/mgt-element)                           | <img src="https://img.shields.io/npm/v/@microsoft/mgt-element/latest.svg">              | <img src="https://img.shields.io/npm/v/@microsoft/mgt-element/next.svg">              |
| [`@microsoft/mgt-components`](https://www.npmjs.com/package/@microsoft/mgt-components)                     | <img src="https://img.shields.io/npm/v/@microsoft/mgt-components/latest.svg">           | <img src="https://img.shields.io/npm/v/@microsoft/mgt-components/next.svg">           |
| [`@microsoft/mgt-react`](https://www.npmjs.com/package/@microsoft/mgt-react)                               | <img src="https://img.shields.io/npm/v/@microsoft/mgt-react/latest.svg">                | <img src="https://img.shields.io/npm/v/@microsoft/mgt-react/next.svg">                |
| [`@microsoft/mgt-msal2-provider`](https://www.npmjs.com/package/@microsoft/mgt-msal2-provider)             | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider/latest.svg">       | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider/next.svg">       |
| [`@microsoft/mgt-teamsfx-provider`](https://www.npmjs.com/package/@microsoft/mgt-teamsfx-provider)         | <img src="https://img.shields.io/npm/v/@microsoft/mgt-teamsfx-provider/latest.svg">     | <img src="https://img.shields.io/npm/v/@microsoft/mgt-teamsfx-provider/next.svg">     |
| [`@microsoft/mgt-sharepoint-provider`](https://www.npmjs.com/package/@microsoft/mgt-sharepoint-provider)   | <img src="https://img.shields.io/npm/v/@microsoft/mgt-sharepoint-provider/latest.svg">  | <img src="https://img.shields.io/npm/v/@microsoft/mgt-sharepoint-provider/next.svg">  |
| [`@microsoft/mgt-proxy-provider`](https://www.npmjs.com/package/@microsoft/mgt-proxy-provider)             | <img src="https://img.shields.io/npm/v/@microsoft/mgt-proxy-provider/latest.svg">       | <img src="https://img.shields.io/npm/v/@microsoft/mgt-proxy-provider/next.svg">       |
| [`@microsoft/mgt-spfx-utils`](https://www.npmjs.com/package/@microsoft/mgt-spfx-utils)                     | <img src="https://img.shields.io/npm/v/@microsoft/mgt-spfx-utils/latest.svg">                 | <img src="https://img.shields.io/npm/v/@microsoft/mgt-spfx/next.svg">                 |
| [`@microsoft/mgt-electron-provider`](https://www.npmjs.com/package/@microsoft/mgt-electron-provider)       | <img src="https://img.shields.io/npm/v/@microsoft/mgt-electron-provider/latest.svg">    | <img src="https://img.shields.io/npm/v/@microsoft/mgt-electron-provider/next.svg">    |

### Preview packages

In addition to the `@next` preview packages, we also ship packages under several other preview tags with various features in progress:

| Tag             | Description                                                              |
| --------------- | ------------------------------------------------------------------------ |
| `next`          | Next release - updated on each commit to `main`                          |

To install these packages, use the tag as the version in your `npm i` command. Ex: `npm i @microsoft/mgt-element@next`. Make sure to install the same version for all mgt packages to avoid any conflicts. Keep in mind, these are features in preview and are not recommended for production use.


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


All web components are also available as React component - see [@microsoft/mgt-react documentation](https://learn.microsoft.com/graph/toolkit/get-started/mgt-react).

## Providers

[Providers](https://learn.microsoft.com/graph/toolkit/providers/providers) enable authentication and provide the implementation for acquiring access tokens on various platforms. The providers also expose a Microsoft Graph Client for calling the Microsoft Graph APIs. The components work best when used with a provider, but the providers can be used on their own as well.

* [Msal2Provider](https://learn.microsoft.com/graph/toolkit/providers/msal2)
* [SharePointProvider](https://learn.microsoft.com/graph/toolkit/providers/sharepoint)
* [TeamsFxProvider](https://learn.microsoft.com/graph/toolkit/providers/teamsfx)
* [ProxyProvider](https://learn.microsoft.com/graph/toolkit/providers/proxy)
* [SimpleProvider](https://learn.microsoft.com/graph/toolkit/providers/custom)
* [ElectronProvider](https://learn.microsoft.com/graph/toolkit/providers/electron)

You can also create your own providers by extending the [IProvider](https://learn.microsoft.com/graph/toolkit/providers/custom) abstract class.

[View the full documentation](https://learn.microsoft.com/graph/toolkit/overview)

## Getting Started

The following guides are available to help you get started with the Toolkit:
* [Build a web application (JavaScript)](https://learn.microsoft.com/graph/toolkit/get-started/build-a-web-app)
* [Build a SharePoint web part Part](https://learn.microsoft.com/graph/toolkit/get-started/build-a-sharepoint-web-part)
* [Build a Microsoft Teams tab](https://learn.microsoft.com/graph/toolkit/get-started/build-a-microsoft-teams-tab)
* [Build an Electron app](https://learn.microsoft.com/en-us/graph/toolkit/get-started/build-an-electron-app)
* [Use the Toolkit with React](https://learn.microsoft.com/graph/toolkit/get-started/use-toolkit-with-react)
* [Use the Toolkit with Angular](https://learn.microsoft.com/graph/toolkit/get-started/use-toolkit-with-angular)
* [Build a productivity hub app](https://learn.microsoft.com/en-us/graph/toolkit/get-started/building-one-productivity-hub)

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

## Using our samples

We, in collaboration with the community, are providing different samples to help you with different scenarios to leverage the Microsoft Graph Toolkit. Our samples are hosted in another repo and is also fully open-source! Head over to the [Microsoft Graph Toolkit Samples Repository](https://aka.ms/mgt/samples) and you will find all sorts of samples to get you started quickly!

## Contribute

We enthusiastically welcome contributions and feedback. Please read our [wiki](https://github.com/microsoftgraph/microsoft-graph-toolkit/wiki) and the [contributing guide](CONTRIBUTING.md) before you begin.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

All files in this GitHub repository are subject to the [MIT license](https://github.com/microsoftgraph/microsoft-graph-toolkit/blob/main/LICENSE). This project also references fonts and icons from a CDN, which are subject to a separate [asset license](https://static2.sharepointonline.com/files/fabric/assets/license.txt).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
