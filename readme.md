# Microsoft Graph Toolkit (Preview)

[![NPM](https://img.shields.io/npm/v/microsoft-graph-toolkit.svg)](https://www.npmjs.com/package/microsoft-graph-toolkit) [![Licence](https://img.shields.io/github/license/microsoftgraph/microsoft-graph-toolkit.svg)](https://github.com/microsoftgraph/msgraph-sdk-javascript) [![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg)](https://github.com/microsoftgraph/msgraph-sdk-javascript) [![stackoverflow](https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg)](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) 
[![Build Status](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_apis/build/status/microsoftgraph.microsoft-graph-toolkit?branchName=master)](https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_build/latest?definitionId=1&branchName=master)

[TODO - add badges for azure pipelines]

The Graph Toolkit is a collection of UI components and helpers for the Microsoft Graph.

Components are built as web components and work with any web framework

## This library is in Preview

This project is in preview and can change at any time. This includes changes in apis, design, and architecture. Based on feedback, the project is expected to evolve and stabilize over time.

## Documentation

[TODO]

## Getting Started

You can install the components by referencing them through our CDN or installing them through NPM

### Install via CDN:

[TODO] setup CDN

```html
<script src="microsoft-graph-toolkit.js"></script>
```

You can then start using the components in your html page

### Install via NPM:

```bash
npm install microsoft-graph-toolkit
```

Then reference `node_modules/microsoft-graph-toolkit/dist/es6/index.js` in the page where you are planning to use it

```html
<script src="node_modules/microsoft-graph-toolkit/dist/es6/index.js"></script>
```

### Usage

The components work best when used with a [provider](./docs/authentication.md). Here is an example of using the [msal provider] (just add anywhere on the page):

```html
<mgt-msal-provider client-id=[CLIENT-ID]></mgt-msal-provider>
```

The toolkit also contains providers for [SharePoint], [Teams], and [Office Add-ins]. You can also create your own providers by implementing the [IProvider] interface.

Any components you add on this page will use the provider to connect to the Microsoft Graph

```html
<mgt-login></mgt-login>

<mgt-person person-query="nikola metulev"></mgt-person>
```

## Contribute

We enthusiastically welcome contributions and feedback. Please read the [contributing guide](CONTRIBUTING.md) before you begin.

## Feedback and Requests
For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

Copyright (c) Microsoft and Contributors. All right reserved. Licensed under the MIT License

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

