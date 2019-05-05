---
title: "Templating the Microsoft Graph Toolkit"
description: "Use custom templates to modify the content of a component
localization_priority: Normal
author: nmetulev
---

# Templates

All web components support templates based on the `<template>` element. For example, to override the template of a component, simply add a `<template>` element inside of a component.

```html
<mgt-agenda>
  <template data-type="event">
      <div>{{event.subject}}</div>
      <div data-for='attendee in event.attendees'>
          <mgt-person person-query="{{attendee.emailAddress.name}}">
            <template>
              <div data-if="person.image">
                <img src="{{person.image}}" />
              </div>
              <div data-else>
                {{person.displayName}}
              </div>
            </template>
          </mgt-person>
      </div>
  </template>
</mgt-agenda>
```

Here you'll notice several template features we support:

- Use the double curly brackets (`{{expression}}`) to expand an expression. In the example above, the `<mgt-person>` passes a `person` object that you can use in the template.
- Use the `data-if` and `data-else` attributes for conditional rendering. Conditional expressions such as `event.attendees.length > 2` are supported
- Use the `data-for` to repeat an element
- Use the `data-type` to specify what part of the component to template. Not specifying the type will template the entire component

Each component documents the supported `data-type` values and what data context is passed down to each template.

The templates can be styled normally as they are rendered outside of the shadow dom.
