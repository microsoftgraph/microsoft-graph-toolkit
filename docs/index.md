# The Microsoft Graph Toolkit

The Microsoft Graph Toolkit is a collection of framework agnostic web components and helpers for accessing and working with the Microsoft Graph. ALl components know how to access the Microsoft Graph out of the box.

## Getting Started

You can install the components by referencing them through our CDN or installing them through NPM

### Use via CDN:

```html
<script src="https://mgtlib.z5.web.core.windows.net/mgt/latest/mgt-loader.js"></script>
```

You can then start using the components in your html page. Here is a full working example with the Msal provider:

```html
<script src="https://mgtlib.z5.web.core.windows.net/mgt/latest/mgt-loader.js"></script>
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
  import { MsalProvider } from '@microsoft/mgt/dist/es6/providers/MsalProvider.js';
  import { Providers } from '@microsoft/mgt/dist/es6/Providers.js';

  Providers.globalProvider = new MsalProvider({ clientId: '[CLIENT-ID]' });
</script>
```

## Providers

The components work best when used with a [provider](./providers.md). The provider exposes authentication and Microsoft Graph apis used by the components to call into the Microsoft Graph.

The toolkit contains providers for [MSAL](./providers/msal.md), [SharePoint](./providers/sharepoint.md), [Teams](./providers/teams.md), and Office Add-ins (coming soon). You can also create your own providers by extending the [IProvider] abstract class.

## Using the components with React, Angular, and other frameworks

Web Components are based on several web standards and can be used with any framework you are already using. However, not all frameworks handle web components the same, and there might be some consideration you should know depending on your framework. The [Custom Elements Everywhere](https://custom-elements-everywhere.com/) project is a great resource for this.

Below is a quick overview of using the microsoft-graph-toolkit components with React and Angular

### React

1. React passes all data to Custom Elements in the form of HTML attributes. For primitive data this is fine, but it does not work when passing rich data, like objects or arrays. In those cases you will need to use a `ref` to pass in the object:

Ex:

```jsx
// import all the components
import '@microsoft/mgt';

class App extends Component {
  render() {
    return <mgt-person show-name ref={el => (el.personDetails = { displayName: 'Nikola Metulev' })} />;
  }
}
```

2. Because React implements its own synthetic event system, it cannot listen for DOM events coming from Custom Elements without the use of a workaround. Developers will need to reference the toolkit components using a ref and manually attach event listeners with addEventListener.

Ex:

```jsx
// you can just import a single component
import '@microsoft/mgt/dist/es6/components/mgt-login/mgt-login.js';

class App extends Component {
  render() {
    return <mgt-login ref="loginComponent" />;
  }

  componentDidMount() {
    this.refs.loginComponent.addEventListener('loginCompleted', e => {
      // handle event
    });
  }
}
```

#### React, Typescript and TSX

There is a known issue using custom elements with React and Typescript where Typescript will throw an error when trying to use a component in tsx. The workaround is to define the custom element in your code:

```ts
declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-login': any;
    }
  }
}
```

You can then use it in your tsx as `<mgt-login></mgt-login>`

### Angular

Angular's default binding syntax will always set properties on an element. This works well for rich data, like objects and arrays, and also works well for primitive values.

To use custom elements, first, enable custom elements in your `app.module.ts` by adding the `CUSTOM_ELEMENT_SCHEMA` to the `@NgModule() decorator`

Ex:

```ts
import { BrowserModule } from '@angular/platform-browser';
import { NgModule, CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';

import { AppComponent } from './app.component';

@NgModule({
  declarations: [AppComponent],
  imports: [BrowserModule],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule {}
```

You can then import the component you'd like to use in your component \*.ts file:

```ts
import { Component } from '@angular/core';
import '@microsoft/mgt/dist/es6/components/mgt-person/mgt-person';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  person = {
    displayName: 'Nikola Metulev'
  };
}
```

And finally, use the component as you normally would in your template

```html
<mgt-person [personDetails]="person" show-name></mgt-person>
```
