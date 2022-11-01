# Microsoft Graph Toolkit

<a href="https://www.npmjs.com/package/@microsoft/mgt"><img src="https://img.shields.io/npm/v/@microsoft/mgt.svg"></a> <a href="https://github.com/microsoftgraph/msgraph-sdk-javascript"><img src="https://cdn.jsdelivr.net/gh/storybookjs/brand@master/badge/badge-storybook.svg"></a>

The Microsoft Graph Toolkit is a collection of reusable, framework-agnostic components and authentication providers for accessing and working with Microsoft Graph. The components are fully functional right of out of the box, with built in providers that authenticate with and fetch data from Microsoft Graph.

The `@microsoft/mgt` package brings all mgt packages together (with the exception of `@microsoft/mgt-react`) and bundles them in this one convenient package.

[View the full documentation](https://learn.microsoft.com/graph/toolkit/overview)

You can now explore components and samples with the [playground](https://mgt.dev).
## Components

The Microsoft Graph Toolkit includes a collection of web components for the most commonly built experiences powered by Microsoft Graph APIs.

The components are also available as [React components](https://learn.microsoft.com/graph/toolkit/get-started/mgt-react).

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

## Providers
[Providers](https://learn.microsoft.com/graph/toolkit/providers) enable authentication and provide the implementation for acquiring access tokens on various platforms and expose a Microsoft Graph Client for calling the Microsoft Graph APIs. The components work best when used with a provider, but the providers can be used on their own.

* [Msal Provider](https://learn.microsoft.com/graph/toolkit/providers/msal)
* [Msal2 Provider](https://learn.microsoft.com/graph/toolkit/providers/msal2)
* [SharePoint Provider](https://learn.microsoft.com/graph/toolkit/providers/sharepoint)
* [Teams Provider](https://learn.microsoft.com/graph/toolkit/providers/teams)
* [Teams Msal2 Provider](https://learn.microsoft.com/graph/toolkit/providers/teams-msal2)
* [Proxy Provider](https://learn.microsoft.com/graph/toolkit/providers/proxy)
* [Simple Provider](https://learn.microsoft.com/graph/toolkit/providers/custom)

## Getting Started

[Watch the Getting Started Video](https://www.youtube.com/watch?v=oZCGb2MMxa0)

You can use the components by referencing the loader directly (via unpkg), or installing the npm package

### Use via mgt-loader:

```html
<script src="https://unpkg.com/@microsoft/mgt@2/dist/bundle/mgt-loader.js"></script>
```

> NOTE: This link will load the highest available version of @microsoft/mgt in the range `>= 2.0.0 < 3.0.0`, omitting the `@2` fragment from the url results in loading the latest version. This could result in loading a new major version and breaking the application.

You can then start using the components in your html page. Here is a full working example with the Msal provider:

```html
<script src="https://unpkg.com/@microsoft/mgt@2/dist/bundle/mgt-loader.js"></script>

<mgt-msal-provider client-id="[CLIENT-ID]"></mgt-msal-provider>

<mgt-login></mgt-login>
<mgt-agenda></mgt-agenda>
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

<mgt-msal-provider client-id="[CLIENT-ID]"></mgt-msal-provider>

<mgt-login></mgt-login>
<mgt-agenda></mgt-agenda>
```

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## Sea also
* [Microsoft Graph Toolkit docs](https://aka.ms/mgt-docs)
* [Microsoft Graph Toolkit repository](https://aka.ms/mgt)
* [Microsoft Graph Toolkit playground](https://mgt.dev)