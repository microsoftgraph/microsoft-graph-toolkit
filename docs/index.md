# The Graph Toolkit

...is a set of web components and helpers for accessing and working with the Microsoft Graph. The library is fully built in the open. 

## Get Started

Download the minified code from [here] and reference the toolkit.js in your html page. 

```html
<script src="./dist/es6/index.js"></script>
```

You can then start using the components in your html page

```html
<mgt-login></mgt-login>

<mgt-person person-id="nikola metulev"></mgt-person>
```

All the components know how to talk to the graph as long as you provide an [authentication context](./authentication.md) they can use.


## NPM
Alternatively, you can install the components using npm

```bash
npm install microsoft-graph-toolkit
```

Then reference `node_modules/microsoft-graph-toolkit/dist/es6/index.js` in the page where you are planning to use it

```html
<script src="node_modules/microsoft-graph-toolkit/dist/es6/index.js"></script>
```

## Testing the preview NPM package
If you want to install the preview package just once, run this command:
```bash
npm install microsoft-graph-toolkit --registry https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
```
If you want to be able to do it again later, by running ```npm install```, add the preview feed to your project's '.npmrc' file (create a '.npmrc' file if you don't have one already):
```
registry=https://pkgs.dev.azure.com/microsoft-graph-toolkit/_packaging/MGT/npm/registry/
always-auth=true
```
Then just install the package:
```bash
npm install microsoft-graph-toolkit
```

TODO: Make sure everybody can access this feed publicly.

## React

// TODO

## Angular

// TODO
