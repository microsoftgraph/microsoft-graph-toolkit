# Style and Templating

The look and feel of components can be customized in two ways.

1. Through custom styles
2. Through `<template>` elements

Let's take a look at each one.

## Styles

It is not possible to style the components as you'd normally would because the component children are hosted in a [shadow dom](https://developer.mozilla.org/en-US/docs/Web/Web_Components/Using_shadow_DOM). Instead, each component documents a set of [CSS custom properties](https://developer.mozilla.org/en-US/docs/Web/CSS/Using_CSS_custom_properties) that can be used to change the look and feel of certain elements.

ex:

```css
mgt-person {
  --avatar-size: 34px;
}
```

## Templates

All web components support templates based on the `<template>` element. For example, to override the template of a component, simply add a `<template>` element inside of a component.

```html
<mgt-person person-query="Nikola">
  <template>
    <div data-if="person.image">
      <img src="{{person.image}}" />
    </div>
    <div data-else>
      {{person.displayName}}
    </div>
  </template>
</mgt-person>
```

```html
<mgt-agenda>
    <template data-type="event">
        <div>{{event.subject}}</div>
        <div data-for="event in events">
            <div data-for='attendee in event.attendees'>
                <mgt-person person-query="{{attendee.emailAddress.name}}">
            </div>
        </div>
    </template>
</mgt-agenda>
```

Here you'll notice several template features we support:

- Use the double curly brackets (`{{expression}}`) to expand an expression. In the example above, the `<mgt-person>` passes a `person` object that you can use in the template.
- Use the `data-if` and `data-else` attributes for conditional rendering
- Use the `data-for` to repeat an element
- Use the `data-type` to specify what part of the component to template. Not specifying the type will template the entire component

Each component documents the supported `data-type` values and what data context is passed down to each template.

The templates can be styled normally as they are rendered outside of the shadow dom.
