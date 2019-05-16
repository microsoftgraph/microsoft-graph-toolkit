---
title: 'Styling the Microsoft Graph Toolkit'
description: 'Use CSS custom properties to modify the component styles'
localization_priority: Normal
author: nmetulev
---

# Styling components

Each component documents a set of [CSS custom properties](https://developer.mozilla.org/en-US/docs/Web/CSS/Using_CSS_custom_properties) that can be used to change the look and feel of certain elements.

ex:

```css
mgt-person {
  --avatar-size: 34px;
}
```

It is not possible to style internal elements of a component unless a CSS custom property is provided. The component children elements are hosted in a [shadow dom](https://developer.mozilla.org/en-US/docs/Web/Web_Components/Using_shadow_DOM).

For more flexibility, consider using [custom templates](./templates.md).

## Addendum:

( Experimental! )

It is now possible to disable the Shadow Dom, and directly style internal elements using normal browser stylesheets, by setting the static property 'useShadowDom' of the MgtBaseComponent class to false before using any MGT Tags:

Example:

```html
<script type="module">
  import { MgtBaseComponent } from './dist/es6/components/baseComponent.js';

  MgtBaseComponent.useShadowDom = false;
</script>

<script type="module" src="./dist/es6/components/mgt-tasks/mgt-tasks.js"></script>
````
