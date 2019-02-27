# Graph Toolkit

> NOTE: This is a very early work in project and everything, including names, apis, design, and architecture can change without notice.

The Graph Toolkit is a collection of UI components and helpers for the Microsoft Graph.

Components are built as web components using Stencil

## Getting Started

Currently, there are two different NPM projects in this repo:

* Toolkit Providers - a set of authentication providers used by the components to access the graph. They enable login to AAD and accessing the graph. This is a dependency for the components to function

* Toolkit Components - web components 

To run the project you need to first build the providers:

```bash
cd providers
npm install
npm run build
```

You can then navigate to the components project and run

```bash
cd components
npm install
npm start
```

To build the component for production, run:

```bash
npm run build
```

To run the unit tests for the components, run:

```bash
npm test
```

## Using this component

### Script tag

- [Publish to NPM](https://docs.npmjs.com/getting-started/publishing-npm-packages)
- Put a script tag similar to this `<script src='https://unpkg.com/my-component@0.0.1/dist/mycomponent.js'></script>` in the head of your index.html
- Then you can use the element anywhere in your template, JSX, html etc

### Node Modules
- Run `npm install my-component --save`
- Put a script tag similar to this `<script src='node_modules/my-component/dist/mycomponent.js'></script>` in the head of your index.html
- Then you can use the element anywhere in your template, JSX, html etc

## About Stencil

Stencil is a compiler for building fast web apps using Web Components.

Stencil combines the best concepts of the most popular frontend frameworks into a compile-time rather than run-time tool.  Stencil takes TypeScript, JSX, a tiny virtual DOM layer, efficient one-way data binding, an asynchronous rendering pipeline (similar to React Fiber), and lazy-loading out of the box, and generates 100% standards-based Web Components that run in any browser supporting the Custom Elements v1 spec.

Stencil components are just Web Components, so they work in any major framework or with no framework at all.

Check out the Stencil docs [here](https://stenciljs.com/docs/my-first-component).