<h1 align="center">
  <img height="120" src="https://github.com/microsoftgraph/microsoft-graph-toolkit/raw/main/assets/graff.png" title="Graff the Giraffe">
  <br>
  Microsoft Graph Toolkit
</h1>

<h4 align="center">Web Components powered by <a href="https://graph.microsoft.com">Microsoft Graph</a></h4>

<p align="center">
  <a href="https://www.npmjs.com/package/@microsoft/mgt"><img src="https://img.shields.io/npm/v/@microsoft/mgt.svg"></a> <a href="https://github.com/microsoftgraph/msgraph-sdk-javascript"><img src="https://img.shields.io/badge/code_style-prettier-ff69b4.svg"></a> <a href="https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit"><img src="https://img.shields.io/stackexchange/stackoverflow/t/microsoft-graph-toolkit.svg"></a>
  <a href="https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_build/latest?definitionId=1&branchName=main"><br><img src="https://dev.azure.com/microsoft-graph-toolkit/microsoft-graph-toolkit/_apis/build/status/microsoftgraph.microsoft-graph-toolkit?branchName=main"></a> <a href="https://www.webcomponents.org/element/@microsoft/mgt"><img src="https://img.shields.io/badge/webcomponents.org-published-blue.svg"></a> <a href="https://mgt.dev"><img src="https://raw.githubusercontent.com/storybooks/brand/main/badge/badge-storybook.svg?sanitize=true"></a>
</p>

<p align="center">
  <iframe src="https://mgt.dev/iframe.html?id=samples-general--login-to-show-agenda"></iframe><br>
  The Microsoft Graph Toolkit is a collection of web components powered by Microsoft Graph. The components are functional, working automatically with Microsoft Graph, and work with any web framework and on all modern browsers.
</p>

<p align="center">
  <a href="#components-&-documentation">Components & Documentation</a> • <a href="#getting-started">Getting Started</a> • <a href="#running-the-samples">Running the Samples</a> 
  • <a href="#contribute">Contribute</a> • <a href="#feedback-and-requests">Feedback & Requests</a> • <a href="#license">License</a> • <a href="#code-of-conduct">Code of Conduct</a>
</p>

## Components & Documentation

You can explore components and samples with the [playground](https://mgt.dev) powered by storybook.

The Toolkit currently includes the following components:

* [mgt-login](https://docs.microsoft.com/graph/toolkit/components/login)
* [mgt-person](https://docs.microsoft.com/graph/toolkit/components/person)
* [mgt-person-card](https://docs.microsoft.com/graph/toolkit/components/person-card)
* [mgt-people](https://docs.microsoft.com/graph/toolkit/components/people)
* [mgt-people-picker](https://docs.microsoft.com/graph/toolkit/components/people-picker)
* [mgt-agenda](https://docs.microsoft.com/graph/toolkit/components/agenda)
* [mgt-tasks](https://docs.microsoft.com/graph/toolkit/components/tasks)
* [mgt-get](https://docs.microsoft.com/graph/toolkit/components/get)
* [mgt-teams-channel-picker](https://docs.microsoft.com/en-us/graph/toolkit/components/teams-channel-picker)

The components work best when used with a [provider](https://docs.microsoft.com/graph/toolkit/providers). The provider handles authentication and the requests to the Microsoft Graph APIs used by the components. The Toolkit contains the following providers:

* [Msal Provider](https://docs.microsoft.com/graph/toolkit/providers/msal)
* [SharePoint Provider](https://docs.microsoft.com/graph/toolkit/providers/sharepoint)
* [Teams Provider](https://docs.microsoft.com/graph/toolkit/providers/teams)
* [Proxy Provider](https://docs.microsoft.com/graph/toolkit/providers/proxy)
* [Simple Provider](https://docs.microsoft.com/graph/toolkit/providers/custom)

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

## Running the samples

Some of our samples are coupled to use the locally built mgt packages instead of the published version from npm. Becuase of this, it's helpful to build the monorepo before attempting to run any of the samples.

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

We enthusiastically welcome contributions and feedback. Please read the [contributing guide](CONTRIBUTING.md) before you begin.

### Code Contribution Challenge
There are many exciting new features or interesting bugs that have been left behind because our team is small with limited capacity. We would love your help! We have tagged these issues with 'help wanted' and/or 'good first issue'. We added them into project board [Community Love](https://github.com/microsoftgraph/microsoft-graph-toolkit/projects/29) for easy tracking. If you see anything you would like to contribute to, you can reach out to [mgt-core team](mgt-core@microsoft.com) or tag one of us ([Beth](https://github.com/beth-panx), [Elise](https://github.com/elisenyang), [Nick](https://github.com/vogtn), [Nikola](https://github.com/nmetulev), and [Shane](https://github.com/shweaver-MSFT)) in the issue for help or further discussion. By submitting a PR to solve an issue mentioned above, you can enter to win some exciting prizes! We have some cool socks and t-shirts waiting for you. Check out the official rules for [Code Contribution Challenge](contest.md). The contest will continue until March 2021 where prizes are awarded every month from now on.

## Feedback and Requests

For general questions and support, please use [Stack Overflow](https://stackoverflow.com/questions/tagged/microsoft-graph-toolkit) where questions should be tagged with `microsoft-graph-toolkit`

Please use [GitHub Issues](https://github.com/microsoftgraph/microsoft-graph-toolkit/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) for bug reports and feature requests. We highly recommend you browse existing issues before opening new issues.

## License

All files in this GitHub repository are subject to the [MIT license](https://github.com/OfficeDev/office-ui-fabric-core/blob/master/LICENSE). This project also references fonts and icons from a CDN, which are subject to a separate [asset license](https://static2.sharepointonline.com/files/fabric/assets/license.txt).

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
