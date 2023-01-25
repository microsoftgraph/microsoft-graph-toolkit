# React wrapper for Microsoft Graph Toolkit

[![npm](https://img.shields.io/npm/v/@microsoft/mgt-react?style=for-the-badge)](https://www.npmjs.com/package/@microsoft/mgt-react)

Use `mgt-react` to simplify usage of [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) web components in React. The library wraps all mgt components and exports them as React components.

## Installation

```bash
npm install @microsoft/mgt-react
```

or

```bash
yarn add @microsoft/mgt-react
```

## Usage

All components are available via the npm package and are named using PascalCase. To use a component, first import it at the top:

```tsx
import { Person } from '@microsoft/mgt-react';
```

You can now use `Person` anywhere in your JSX as a regular React component.

```tsx
<Person personQuery="me" />
```

All properties and events map exactly as they are defined in the component documentation - [see web component docs](https://aka.ms/mgt-docs).

For example, you can set the `personDetails` property to an object:

```jsx
const App = (props) => {
  const personDetails = {
    displayName: 'Bill Gates',
  };

  return <Person personDetails={personDetails}></Person>;
};
```

Or, register an event handler:

```jsx
import { PeoplePicker, People } from '@microsoft/mgt-react';

const App = (props) => {
  const [people, setPeople] = useState([]);

  const handleSelectionChanged = (e) => {
    setPeople(e.target.selectedPeople);
  };

  return
    <div>
      <PeoplePicker selectionChanged={handleSelectionChanged} />
      Selected People: <People people={people} />
    </div>;
};
```

## Templates

Most Microsoft Graph Toolkit components [support templating](https://learn.microsoft.com/graph/toolkit/customize-components/templates) and `mgt-react` allows you to leverage React for writing templates.

For example, to create a template to be used for rendering events in the `mgt-agenda` component, first define a component to be used for rendering an event:

```tsx
import { MgtTemplateProps } from '@microsoft/mgt-react';

const MyEvent = (props: MgtTemplateProps) => {
  const { event } = props.dataContext;
  return <div>{event.subject}</div>;
};
```

Then use it as a child of the wrapped component and set the template prop to `event`

```tsx
import { Agenda } from '@microsoft/mgt-react';

const App = (props) => {
  return <Agenda>
    <MyEvent template="event">
  </Agenda>
}
```

The `template` prop allows you to specify which template to overwrite. In this case, the `MyEvent` component will be repeated for every event, and the `event` object will be passed as part of the `dataContext` prop.

## Why

If you've used web components in React, you know that proper interop between web components and React components requires a bit of extra work.

From [https://custom-elements-everywhere.com/](https://custom-elements-everywhere.com/):

> React passes all data to Custom Elements in the form of HTML attributes. For primitive data this is fine, but the system breaks down when passing rich data, like objects or arrays. In these instances you end up with stringified values like some-attr="[object Object]" which can't actually be used.

> Because React implements its own synthetic event system, it cannot listen for DOM events coming from Custom Elements without the use of a workaround. Developers will need to reference their Custom Elements using a ref and manually attach event listeners with addEventListener. This makes working with Custom Elements cumbersome.
