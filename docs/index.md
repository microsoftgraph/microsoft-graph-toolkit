# The Microsoft Graph Toolkit

...is a set of web components and helpers for accessing and working with the Microsoft Graph. The library is fully built in the open. 

## Get Started

Download the minified code from [here] and reference the toolkit.js in your html page. 

```html
<script src="./dist/es6/index.js"></script>
```

You can then start using the components in your html page

```html
<mgt-login></mgt-login>

<mgt-person person-query="nikola metulev"></mgt-person>
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

## Using the components with React, Angular, and other frameworks

Web Components are based on several web standards and can be used with any framework you are already using. However, not all frameworks handle web components the same, and there might be some consideration you should know depending on your framework. The [Custom Elements Everywhere](https://custom-elements-everywhere.com/) project is a great resource for this.

Below is a quick overview of using the microsoft-graph-toolkit components with React and Angular

### React

1. React passes all data to Custom Elements in the form of HTML attributes. For primitive data this is fine, but it does not work when passing rich data, like objects or arrays. In those cases you will need to use a `ref` to pass in the object:

Ex:

```jsx
// import all the components
import "microsoft-graph-toolkit";

class App extends Component {
    render() {
        return (
            <mgt-person show-name ref={el => el.personDetails = {displayName: 'Nikola Metulev'}}></mgt-person>
        );
    }
}
```

2. Because React implements its own synthetic event system, it cannot listen for DOM events coming from Custom Elements without the use of a workaround. Developers will need to reference the toolkit components using a ref and manually attach event listeners with addEventListener.

Ex:
```jsx
// you can just import a single component
import "microsoft-graph-toolkit/dist/es6/components/mgt-login/mgt-login.js";

class App extends Component {
  render() {
    return (
        <mgt-login ref="loginComponent"></mgt-login>
    );
  }

  componentDidMount() {
    this.refs.loginComponent.addEventListener('loginCompleted', e => {
        // handle event
    });
  }
}

```

### Angular

Angular's default binding syntax will always set properties on an element. This works well for rich data, like objects and arrays, and also works well for primitive values.

To use custom elements, first, enable custom elements in your `app.module.ts` by adding the `CUSTOM_ELEMENT_SCHEMA` to the `@NgModule() decorator`

Ex:

```ts
import { BrowserModule } from '@angular/platform-browser';
import { NgModule, CUSTOM_ELEMENTS_SCHEMA } from '@angular/core';

import { AppComponent } from './app.component';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule {
}

```

You can then import the component you'd like to use in your component *.ts file:

```ts
import { Component } from '@angular/core';
import 'microsoft-graph-toolkit/dist/es6/components/mgt-person/mgt-person';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  person = {
    displayName: "Nikola Metulev"
  };
}
```

And finally, use the component as you normally would in your template

```html
<mgt-person [personDetails]="person" show-name></mgt-person>
```

