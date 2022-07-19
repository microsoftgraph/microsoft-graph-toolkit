# DemoMgtAngular

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 9.1.5.

## Getting Started
The is the sample code explained in [A Lap Around the Mircosoft Graph Toolkit Day 14 - How to use Microsoft Graph Toolkit with Angular](https://developer.microsoft.com/en-us/graph/blogs/a-lap-around-microsoft-graph-toolkit-day-14-using-microsoft-graph-toolkit-with-angular/)

Documentation for Msal2Provider can be found here (https://docs.microsoft.com/en-us/graph/toolkit/get-started/use-toolkit-with-angular)

> NOTE: This sample is linked to the locally built mgt packages. If you haven't already, run the following commands in the root of the repo before getting started.
```bash
npm i -g yarn
yarn
yarn build
```

1. In a web browser, navigate to the Azure Portal and create a new [App Registration](http://aka.ms/AppRegistrations) with the following Graph API permissions:
    * `openid`
    * `profile`
    * `user.read`
    * `calendars.read`
1. Record the **Client Id** value for later.
1. Modify `src/environment/environment.msal.ts` by replacing `[YOUR-CLIENT-ID]` with the Azure App Registration **Client Id** from earlier.
1. Run `yarn start` for a local dev server and navigate to `http://localhost:4200/`. The app will automatically reload if you change any of the source files.
