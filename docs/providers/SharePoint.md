# SharePoint provider

Use the SharePoint provider inside your SharePoint web parts to power the components with Microsoft Graph access.

Visit [the authentication docs](../providers.md) to learn more about the role of providers in the Microsoft Graph Toolkit

## Getting started

Initialize the provider inside of your `onInit()` method of your web part.

```ts

// import the providers at the top of the page
import {Providers, SharePointProvider} from '@microsoft/mgt';

// add the onInit() method if not already there in your web part class
protected async onInit() {
    Providers.globalProvider = new SharePointProvider(this.context);
}
```

Now you can add any component in your `render()` method and it will use the SharePoint context to access the Microsoft Graph.

```ts

public render(): void {
    this.domElement.innerHTML = `
      <mgt-agenda></mgt-agenda>
      `;
  }
```

> Note: The Microsoft Graph Toolkit requires Typescript 3.x. Make sure you are using a supported version of Typescript by [installing the right compiler](https://github.com/SharePoint/sp-dev-docs/wiki/SharePoint-Framework-v1.8-release-notes#support-for-typescript-27-29-and-3x).

## Testing in the workbench

If you are just getting started with SharePoint web parts, you can follow the getting started guide [here](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part)

Once you have a web part, and you are ready to use the components, you will need to make sure your web part has the right permissions to access the graph. [This tutorial](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aad-tutorial) does a great job at walking you through the process.

In a nutshell, it's important to add the right permission to your `package-solution.json` as described in the linked tutorial. You will need to upload a package of your web part to SharePoint and have an Administrator approve the requested permissions.

> hint: if you are not sure what permissions to add, each component documentation includes all the permissions it needs.
