## SharePoint webpart using Microsoft Graph Toolkit components

This sample is a React based SharePoint webpart that uses the Microsoft Graph Toolkit.

### Building the code

```bash
> git clone https://github.com/microsoft/microsoft-graph-toolkit

> cd microsoft-graph-toolkit

> npm i -g gulp yarn #if you don't have them installed globally

> yarn

> yarn build

> cd samples/sp-webpart

> gulp trust-dev-cert

> gulp
```

### Testing the webpart

Run the webpart by running:

`gulp serve`

Open the browser to:

`https://<your-tenant>.sharepoint.com/_layouts/15/Workbench.aspx`

The webpart will be available to be added

### Approving permissions 

The admin of your tenant will need to approve the graph permissions in the admin portal before you can use any graph APIs. You only need to do these steps once (or if you add additional permissions to `config/package-solution.json`)

1. Package the web part. First run the following commands:

    ```
    gulp build
    gulp bundle
    gulp package-solution
    ```

    This will create a package under `sharepoint/solution`, with the `.sppkg` extension

2. Upload the package to an app catalog. Navigate to the app catalog url and click on *Distribute apps for SharePoint*. Click on *new* to upload the solution to the app catalog and deploy it.

    > If you do not have an app catalog on your tenant, follow [these steps](https://docs.microsoft.com/en-us/sharepoint/use-app-catalog#step-1-create-the-app-catalog-site-collection) to create one. 

3. Approve the permissions. Go to the SharePoint admin center (best reached via https://admin.microsoft.com > click **Show All** > SharePoint) and make sure you are signed in as admin.

    Click on the *Advanced* menu, you should see *API access* which after clicking will show the page with permissions that need to be approved. Approve each one by one.

You should now be able to test the webpart

### Build options

gulp trust-dev-cert

gulp clean
gulp test
gulp bundle
gulp package-solution
