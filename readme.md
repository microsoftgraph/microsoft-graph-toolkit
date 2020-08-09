# Microsoft Graph Toolkit

<img align="left" height="150" src="https://github.com/microsoftgraph/microsoft-graph-toolkit/raw/main/assets/graff.png" title="Graff the Giraffe">

[![NPM](https://img.shields.io/npm/v/@microsoft/mgt.svg)](https://www.npmjs.com/package/@microsoft/mgt) [![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg)](https://github.com/microsoftgraph/msgraph-sdk-javascript) [![stackoverflow](https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg)](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit)
[![Build Status](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_apis/build/status/microsoftgraph.microsoft-graph-toolkit?branchName=main)](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_build/latest?definitionId=1&branchName=main) [![Published on webcomponents.org](https://img.shields.io/badge/webcomponents.org-published-blue.svg)](https://www.webcomponents.org/element/@microsoft/mgt) [![Storybook](https://raw.githubusercontent.com/storybooks/brand/main/badge/badge-storybook.svg?sanitize=true)](https://mgt.dev)

The Microsoft Graph Toolkit is a collection of web components powered by the Microsoft Graph.

Components are functional and work automatically with the Microsoft Graph

Components work with any web framework and on all modern browsers. IE 11 is also supported

[Here is a quick jsfiddle](https://jsfiddle.net/metulev/9phqxLd5/)

## Components & Documentation

The toolkit currently includes the following components:

* [mgt-login](https://docs.microsoft.com/graph/toolkit/components/login)
* [mgt-person](https://docs.microsoft.com/graph/toolkit/components/person)
* [mgt-person-card](https://docs.microsoft.com/graph/toolkit/components/person-card)
* [mgt-people](https://docs.microsoft.com/graph/toolkit/components/people)
* [mgt-people-picker](https://docs.microsoft.com/graph/toolkit/components/people-picker)
* [mgt-agenda](https://docs.microsoft.com/graph/toolkit/components/agenda)
* [mgt-tasks](https://docs.microsoft.com/graph/toolkit/components/tasks)
* [mgt-get](https://docs.microsoft.com/graph/toolkit/components/get)
* [mgt-teams-channel-picker](https://docs.microsoft.com/en-us/graph/toolkit/components/teams-channel-picker)

And the following providers:

* [Msal Provider](https://docs.microsoft.com/graph/toolkit/providers/msal)
* [SharePoint Provider](https://docs.microsoft.com/graph/toolkit/providers/sharepoint)
* [Teams Provider](https://docs.microsoft.com/graph/toolkit/providers/teams)
* [Proxy Provider](https://docs.microsoft.com/graph/toolkit/providers/proxy)
* [Simple Provider](https://docs.microsoft.com/graph/toolkit/providers/custom)

[View the full documentation](https://docs.microsoft.com/graph/toolkit/overview)

You can now explore components and samples with the [playground](https://mgt.dev) powered by storybook.

## Getting Started

[Watch the Getting Started Video](https://www.youtube.com/watch?v=oZCGb2MMxa0)

You can use the components by referencing the loader directly (via unpkg), or installing the npm package

### Use via mgt-loader:

```html
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
```

You can then start using the components in your html page. Here is a full working example with the Msal provider:

```html
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
<mgt-msal-provider client-id="[CLIENT-ID]"></mgt-msal-provider>
<mgt-login></mgt-login>

<!-- <script>
    // alternatively, you can set the provider in code and provide more options
    mgt.Providers.globalProvider = new mgt.MsalProvider({clientId: '[CLIENT-ID]'});
</script> -->
```

> NOTE: MSAL requires the page to be hosted in a web server for the authentication redirects. If you are just getting started and want to play around, the quickest way is to use something like [live server](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer) in vscode.

### Use via NPM:

The benefits of using MGT through NPM is that you have full control of the bundling process and you can bundle only the code you need for your site. First, add the npm package:

```bash
npm install @microsoft/mgt
```

Now you can reference all components at the page you are using:

```html
<script src="node_modules/@microsoft/mgt/dist/es6/components.js"></script>
```

Or, just reference the component you need and avoid loading everything else:

```html
<script src="node_modules/@microsoft/mgt/dist/es6/components/mgt-login/mgt-login.js"></script>
```

Similarly, to add a provider, you can add it as a component:

```html
<script src="node_modules/@microsoft/mgt/dist/es6/components/providers/mgt-msal-provider.js"></script>

<mgt-msal-provider client-id="[CLIENT-ID]"></mgt-msal-provider>
```

or, add it in your code:

```html
<script type="module">
  import { Providers, MsalProvider } from '@microsoft/mgt';

  Providers.globalProvider = new MsalProvider({ clientId: '[CLIENT-ID]' });
</script>
```

or, if you already use MSAL.js and have a `UserAgentApplication`, you can use it:

```html
<script type="module">
  import { Providers, MsalProvider } from '@microsoft/mgt';

  const app = new UserAgentApplication({ auth: { clientId: '[CLIENT-ID]' } });

  Providers.globalProvider = new MsalProvider({ userAgentApplication: app });
</script>
```

## Providers

The components work best when used with a [provider](https://docs.microsoft.com/graph/toolkit/providers). The provider exposes authentication and Microsoft Graph apis used by the components to call into the Microsoft Graph.

The toolkit contains providers for [MSAL](https://docs.microsoft.com/graph/toolkit/providers/msal), [SharePoint](https://docs.microsoft.com/graph/toolkit/providers/sharepoint), and [Teams](https://docs.microsoft.com/graph/toolkit/providers/teams). You can also create your own providers by extending the [IProvider](https://docs.microsoft.com/graph/toolkit/providers/custom) abstract class.

## Contribute

We enthusiastically welcome contributions and feedback. Please read the [contributing guide](CONTRIBUTING.md) before you begin.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

All files in this GitHub repository are subject to the [MIT license](https://github.com/OfficeDev/office-ui-fabric-core/blob/master/LICENSE). This project also references fonts and icons from a CDN, which are subject to a separate [asset license](https://static2.sharepointonline.com/files/fabric/assets/license.txt).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
