# SharePoint Framework library for Microsoft Graph Toolkit

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-spfx?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-spfx)

![SPFx 1.11](https://img.shields.io/badge/SPFx-1.11-green.svg?style=for-the-badge)

Use the SharePoint Framework library for Microsoft Graph Toolkit to use Microsoft Graph Toolkit in SharePoint Framework solutions.

To prevent multiple components from registering their own set of Microsoft Graph Toolkit components on the page, you should deploy this library to your tenant and reference Microsoft Graph Toolkit components that you use in your solution from this library.

## Installation

To load Microsoft Graph Toolkit components from the library, add the `@microsoft/mgt-spfx` package as a runtime dependency to your SharePoint Framework project:

```bash
npm install @microsoft/mgt-spfx
```

or

```bash
yarn add @microsoft/mgt-spfx
```

Before deploying your SharePoint Framework package to your tenant, you will need to deploy the `@microsoft/mgt-spfx` SharePoint Framework package to your tenant. You can download the package corresponding to the version of `@microsoft/mgt-spfx` that you used in your project, from the [Releases](https://github.com/microsoftgraph/microsoft-graph-toolkit/releases) section on GitHub.

**Important:** Since there can be only one version of the SharePoint Framework library for Microsoft Graph Toolkit installed in the tenant, before using MGT in your solution, consult with your organization/customer if they already have a version of SharePoint Framework library for Microsoft Graph Toolkit deployed in their tenant and use the same version to avoid issues.

## Usage

When building SharePoint Framework web parts and extensions, reference the Microsoft Graph Toolkit `Provider` and `SharePointProvider` from the `@microsoft/mgt-spfx` package. This will ensure, that your solution will use MGT components that are already registered on the page, rather than instantiating its own:

```ts
import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

// [...] trimmed for brevity

export default class MgtReactWebPart extends BaseClientSideWebPart<IMgtReactWebPartProps> {
  protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
  }

  // [...] trimmed for brevity
}
```

When referencing the actual components, you can load them directly from the `@microsoft/mgt`, or the `@microsoft/mgt-react` package if you use React.

```tsx
import { Person } from '@microsoft/mgt-react';

// [...] trimmed for brevity

export default class MgtReact extends React.Component<IMgtReactProps, {}> {
  public render(): React.ReactElement<IMgtReactProps> {
    return (
      <div className={ styles.mgtReact }>
        <Person personQuery="me" />
      </div>
    );
  }
}
```
