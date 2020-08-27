# DemoMgtAngular

This project was generated with [Angular CLI](https://github.com/angular/angular-cli) version 9.1.5.

## Getting Started
The is the sample code explained in [A Lap Around the Mircosoft Graph Toolkit Day 14 - How to use Microsoft Graph Toolkit with Angular](https://developer.microsoft.com/en-us/graph/blogs/a-lap-around-microsoft-graph-toolkit-day-14-using-microsoft-graph-toolkit-with-angular/)

1. Run `npm install` (to download all dependant npm packages)
2. Create an Azure App Registration with the following Graph API permissions
    * openid
    * profile
    * user.read
    * calendars.read
2. Modify src/app/app.component.ts and replace [YOUR-CLIENT-ID] with an Azure App Registration client id
3. Run `ng serve` for a local dev server and navigate to `http://localhost:4200/`. The app will automatically reload if you change any of the source files.
