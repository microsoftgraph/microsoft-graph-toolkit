# Microsoft Graph Toolkit (Preview)

<img align="left" height="150" src="https://github.com/microsoftgraph/microsoft-graph-toolkit/raw/master/assets/graff.png" title="Graff the Giraffe">

[![NPM](https://img.shields.io/npm/v/@microsoft/mgt.svg)](https://www.npmjs.com/package/@microsoft/mgt) [![Licence](https://img.shields.io/github/license/microsoftgraph/microsoft-graph-toolkit.svg)](https://github.com/microsoftgraph/msgraph-sdk-javascript) [![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg)](https://github.com/microsoftgraph/msgraph-sdk-javascript) [![stackoverflow](https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg)](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit)
[![Build Status](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_apis/build/status/microsoftgraph.microsoft-graph-toolkit?branchName=master)](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_build/latest?definitionId=1&branchName=master)

The Microsoft Graph Toolkit is a collection of web components powered by the Microsoft Graph.

Components are functional and work automatically with the Microsoft Graph

Components work with any web framework and on all all modern browsers. IE 11 is also supported

[Here is a quick jsfiddle](https://jsfiddle.net/metulev/9phqxLd5/)

## This library is in Preview

This library is in preview and is in early development. Based on feedback from the community, all components and APIs are expected to change and improve.

## Documentation

[View the documentation](./docs)

## Getting Started

[![Watch the Getting Started Video](https://img.youtube.com/vi/oZCGb2MMxa0/2.jpg)](https://www.youtube.com/watch?v=oZCGb2MMxa0)

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

## Providers

The components work best when used with a [provider](./docs/providers.md). The provider exposes authentication and Microsoft Graph apis used by the components to call into the Microsoft Graph.

The toolkit contains providers for [MSAL](./docs/providers/msal.md), [SharePoint](./docs/providers/sharepoint.md), [Teams](./docs/providers/teams.md), and Office Add-ins (coming soon). You can also create your own providers by extending the [IProvider] abstract class.

## Contribute

We enthusiastically welcome contributions and feedback. Please read the [contributing guide](CONTRIBUTING.md) before you begin.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

Copyright (c) Microsoft and Contributors. All right reserved. Licensed under the MIT License

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
