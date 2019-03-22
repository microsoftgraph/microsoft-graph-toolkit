# Graph Toolkit

> NOTE: This is a very early work in project and everything, including names, apis, design, and architecture can change without notice.

The Graph Toolkit is a collection of UI components and helpers for the Microsoft Graph.

Components are built as web components using lit-element

## Getting Started

To run the project, run the following in the terminal/cmd

```bash
npm install
npm start
```

Take a look at `index.html` on how to use these components

## Get Started

Download the minified code from [here] and reference the  in your html page. 

```html
<script src="microsoft-graph-toolkit.js"></script>
```

You can then start using the components in your html page

```html
<mgt-login></mgt-login>

<mgt-person person-query="nikola metulev"></mgt-person>
```

All the components know how to talk to the graph as long as you provide an [authentication context](./docs/authentication.md) they can use.


## NPM
Alternatively, you can install the components using npm

```bash
npm install microsoft-graph-toolkit
```

Then reference `node_modules/microsoft-graph-toolkit/dist/es6/index.js` in the page where you are planning to use it

```html
<script src="node_modules/microsoft-graph-toolkit/dist/es6/index.js"></script>
```