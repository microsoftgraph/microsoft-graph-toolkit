# Vue.js Sample for MGT

This sample shows how to incorporate the Microsoft Graph Toolkit into a [Vue.js](https://vuejs.org/) app.

The sample was initially created with the [vue-cli](https://cli.vuejs.org/) tool for Typescript and dart-sass.

## Setting up Vue for MGT

1. Install MGT with `yarn add @microsoft/mgt @microsoft/mgt-element`
2. In your `App.vue` or other Vue components, add an import with the components you want to use E.g. `import { MgtMsalProvider, MgtLogin, MgtAgenda } from '@microsoft/mgt';`.
3. Now you can use those MGT components in your Vue component!

## Project setup
```bash
# Run these commands at the repo root directory to build the local mgt packages
# and install dependencies in node_modules folders.
npm i -g yarn
yarn
yarn build
cd ./samples/vue-app
```

### Compiles and hot-reloads for development
```bash
yarn serve
```

### Compiles and minifies for production
```bash
yarn build
```

### Run your tests
```bash
yarn test
```

### Lints and fixes files
```
yarn lint
```

### Customize configuration
See [Configuration Reference](https://cli.vuejs.org/config/).
