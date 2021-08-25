<h1 align="center">
  <img height="120" src="https://github.com/microsoftgraph/microsoft-graph-toolkit/raw/main/assets/graff.png" title="Graff the Giraffe">
  <br>
  Microsoft Graph Toolkit
</h1>

<h4 align="center">UI Components and Authentication Providers for <a href="https://graph.microsoft.com">Microsoft Graph</a></h4>

<p align="center">
  <a href="https://github.com/microsoftgraph/msgraph-sdk-javascript"><img src="https://img.shields.io/badge/code_style-prettier-ff69b4.svg"></a> <a href="https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit"><img src="https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg"></a>
  <img src="https://github.com/microsoftgraph/microsoft-graph-toolkit/workflows/Build%20CI/badge.svg" /> <a href="https://www.webcomponents.org/element/@microsoft/mgt"><img src="https://img.shields.io/badge/webcomponents.org-published-blue.svg"></a> <a href="https://mgt.dev"><img src="https://cdn.jsdelivr.net/gh/storybookjs/brand@master/badge/badge-storybook.svg"></a> <a href="https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+label%3A%22good+first+issue%22"><img src="https://img.shields.io/github/issues/microsoftgraph/microsoft-graph-toolkit/good%20first%20issue?color=brightgreen"></a>
</p>

<h3 align="center"><a href="https://aka.ms/mgt-docs">Documentation</a></h3>

<p align="center">
  The Microsoft Graph Toolkit is a collection of reusable, framework-agnostic components and authentication providers for accessing and working with Microsoft Graph. The components are fully functional right of out of the box, with built in providers that authenticate with and fetch data from Microsoft Graph.
</p>

<p align="center">
  <a href="#packages">Packages</a> • <a href="#components">Components</a> • <a href="#providers">Providers</a> • <a href="#getting-started">Getting Started</a> • <a href="#running-the-samples">Running the Samples</a> • <a href="#contribute">Contribute</a> • <a href="#feedback-and-requests">Feedback & Requests</a> <br>• <a href="#license">License</a> • <a href="#code-of-conduct">Code of Conduct</a>
</p>

## Packages

| Package | Latest | Next |
| - | - | - |
| [`@microsoft/mgt`](https://www.npmjs.com/package/@microsoft/mgt) | <img src="https://img.shields.io/npm/v/@microsoft/mgt/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt/next.svg">
| [`@microsoft/mgt-element`](https://www.npmjs.com/package/@microsoft/mgt-element) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-element/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-element/next.svg"> |
| [`@microsoft/mgt-components`](https://www.npmjs.com/package/@microsoft/mgt-components) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-components/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-components/next.svg"> |
| [`@microsoft/mgt-react`](https://www.npmjs.com/package/@microsoft/mgt-react) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-react/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-react/next.svg"> |
| [`@microsoft/mgt-msal-provider`](https://www.npmjs.com/package/@microsoft/mgt-msal-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal-provider/next.svg"> |
| [`@microsoft/mgt-msal2-provider`](https://www.npmjs.com/package/@microsoft/mgt-msal2-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-msal2-provider/next.svg"> |
| [`@microsoft/mgt-teams-provider`](https://www.npmjs.com/package/@microsoft/mgt-teams-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-teams-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-teams-provider/next.svg"> |
| [`@microsoft/mgt-sharepoint-provider`](https://www.npmjs.com/package/@microsoft/mgt-sharepoint-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-sharepoint-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-sharepoint-provider/next.svg"> |
| [`@microsoft/mgt-proxy-provider`](https://www.npmjs.com/package/@microsoft/mgt-proxy-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-proxy-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-proxy-provider/next.svg"> |
| [`@microsoft/mgt-spfx`](https://www.npmjs.com/package/@microsoft/mgt-spfx) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-spfx/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-spfx/next.svg"> |
| [`@microsoft/mgt-electron-provider`](https://www.npmjs.com/package/@microsoft/mgt-electron-provider) | <img src="https://img.shields.io/npm/v/@microsoft/mgt-electron-provider/latest.svg"> | <img src="https://img.shields.io/npm/v/@microsoft/mgt-electron-provider/next.svg"> |


### Preview packages

In addition to the `@next` preview packages, we also ship packages under several other preview tags with various features in progress:

| Tag | Description |
| - | - |
| `next` | Next release - updated on each commit to `main` |
| `next.fluentui` | Next major release (v3) with components based on FluentUI web components |

To install these packages, use the tag as the version in your `npm i` command. Ex: `npm i @microsoft/mgt-element@next.fluentui`. Make sure to install the same version for all mgt packages to avoid any conflicts. Keep in mind, these are features in preview and are not recommended for production use.


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
* [mgt-teams-channel-picker](https://docs.microsoft.com/graph/toolkit/components/teams-channel-picker)
* [mgt-file](https://docs.microsoft.com/graph/toolkit/components/file)
* [mgt-file-list](https://docs.microsoft.com/graph/toolkit/components/file-list)
* [mgt-get](https://docs.microsoft.com/graph/toolkit/components/get)

All web components are also available as React component - see [@microsoft/mgt-react documentation](https://docs.microsoft.com/graph/toolkit/get-started/mgt-react).

## Providers

[Providers](https://docs.microsoft.com/graph/toolkit/providers) enable authentication and provide the implementation for acquiring access tokens on various platforms. The providers also expose a Microsoft Graph Client for calling the Microsoft Graph APIs. The components work best when used with a provider, but the providers can be used on their own as well.

* [Msal Provider](https://docs.microsoft.com/graph/toolkit/providers/msal)
* [Msal2 Provider](https://docs.microsoft.com/graph/toolkit/providers/msal2)
* [SharePoint Provider](https://docs.microsoft.com/graph/toolkit/providers/sharepoint)
* [Teams Provider](https://docs.microsoft.com/graph/toolkit/providers/teams)
* [Proxy Provider](https://docs.microsoft.com/graph/toolkit/providers/proxy)
* [Simple Provider](https://docs.microsoft.com/graph/toolkit/providers/custom)
* [Electron Provider](https://docs.microsoft.com/graph/toolkit/providers/electron)

You can also create your own providers by extending the [IProvider](https://docs.microsoft.com/graph/toolkit/providers/custom) abstract class.

[View the full documentation](https://docs.microsoft.com/graph/toolkit/overview)

## Getting Started

The following guides are available to help you get started with the Toolkit:
* [Build a web application (vanilla JS)](https://docs.microsoft.com/graph/toolkit/get-started/build-a-web-app)
* [Build a SharePoint Web Part](https://docs.microsoft.com/graph/toolkit/get-started/build-a-sharepoint-web-part)
* [Build a Microsoft Teams Tab](https://docs.microsoft.com/graph/toolkit/get-started/build-a-microsoft-teams-tab)
* [Use the Toolkit with React](https://docs.microsoft.com/graph/toolkit/get-started/use-toolkit-with-react)
* [Use the Toolkit with Angular](https://docs.microsoft.com/graph/toolkit/get-started/use-toolkit-with-angular)

You can use the components by referencing the loader directly (via unpkg), or installing the npm package

### Use via mgt-loader:

```html
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
```

You can then start using the components in your html page. Here is a full working example with the Msal2 provider:

```html
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
<mgt-msal2-provider client-id="[CLIENT-ID]"></mgt-msal2-provider>
<mgt-login></mgt-login>

<!-- <script>
    // alternatively, you can set the provider in code and provide more options
    mgt.Providers.globalProvider = new mgt.Msal2Provider({clientId: '[CLIENT-ID]'});
</script> -->
```

> NOTE: MSAL requires the page to be hosted in a web server for the authentication redirects. If you are just getting started and want to play around, the quickest way is to use something like [live server](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer) in vscode.

### Use via NPM:

The benefits of using MGT through NPM is that you have full control of the bundling process and you can bundle only the code you need for your site. First, add the npm package:

```bash
npm install @microsoft/mgt
```

Now you can reference all components and providers at the page you are using:

```html
<script type="module" src="node_modules/@microsoft/mgt/dist/es6/index.js"></script>

<mgt-msal2-provider client-id="[CLIENT-ID]"></mgt-msal2-provider>

<mgt-login></mgt-login>
<mgt-agenda></mgt-agenda>
```

## Running the samples

Some of our samples are coupled to use the locally built mgt packages instead of the published version from npm. Because of this, it's helpful to build the monorepo before attempting to run any of the samples.

```bash
# Starting at the root
yarn
yarn build
# Now you can run the React sample using the local packages
cd ./samples/react-app/
yarn start
```

This also means that running the samples in isolation may fail if there are breaking changes between the published version of mgt and the local copy. 
To workaround this, use samples that are known to be compatible with a specific release by checking out the appropriate branch or tag first. 

## Contribute

We enthusiastically welcome contributions and feedback. Please read our [wiki](https://github.com/microsoftgraph/microsoft-graph-toolkit/wiki) and the [contributing guide](CONTRIBUTING.md) before you begin.

### Code Contribution Challenge
There are many exciting new features or interesting bugs that have been left behind because our team is small with limited capacity. We would love your help! We have tagged these issues with 'help wanted' and/or 'good first issue'. We added them into project board [Community Love](https://github.com/microsoftgraph/microsoft-graph-toolkit/projects/29) for easy tracking. If you see anything you would like to contribute to, you can reach out to  mgt-help@microsoft.com or tag one of us ([Beth](https://github.com/beth-panx), [Elise](https://github.com/elisenyang), [Nick](https://github.com/vogtn), [Nikola](https://github.com/nmetulev), and [Shane](https://github.com/shweaver-MSFT)) in the issue for help or further discussion. By submitting a PR to solve an issue mentioned above, you can enter to win some exciting prizes! We have some cool socks and t-shirts waiting for you. Check out the official rules for [Code Contribution Challenge](contest.md). The contest will continue until March 2021 where prizes are awarded every month from now on.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

All files in this GitHub repository are subject to the [MIT license](https://github.com/OfficeDev/office-ui-fabric-core/blob/master/LICENSE). This project also references fonts and icons from a CDN, which are subject to a separate [asset license](https://static2.sharepointonline.com/files/fabric/assets/license.txt).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
