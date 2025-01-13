/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-theme-toggle / HTML',
  component: 'theme-toggle',
  decorators: [
    withCodeEditor({
      disableThemeToggle: true
    })
  ]
};

export const userPreferenceDriven = () => html`
  <mgt-theme-toggle></mgt-theme-toggle>
  <p>
    This toggle will set light or dark mode based on the user's preference set in the browser.
  </p>
  <style>
body {
    background-color: var(--fill-color);
    color: var(--neutral-foreground-rest);
    font-family: var(--body-font);
}
</style>
`;

export const darkModeOn = () => html`
  <mgt-theme-toggle mode="dark"></mgt-theme-toggle>
  <style>
body {
    background-color: var(--fill-color);
}
  </style>
`;

export const lightModeOn = () => html`
  <mgt-theme-toggle mode="light"></mgt-theme-toggle>
  <style>
body {
    background-color: var(--fill-color);
}
  </style>
`;

export const themingWithoutToggle = () => html`
<h2>Style the page</h2>
<p>
  This page has the provided light theme applied by calling <code>applyTheme('light')</code>.
  and applying some basic css rules to the body tag<br/>
  This function uses a default value of document for the second optional element parameter,
  which is used below to set the dark theme.<br/>
  Refer to the CSS and JS tabs for details.
</p>

<h2>Style the background of a component</h2>
<p>
  The first login component is in the default light theme on a light
  background. We are setting the background of the component with id "login-two"
  to the <code>aliceblue</code> color using CSS.<br/>
  Check the CSS tab.
</p>
<mgt-login id="login-one"></mgt-login>
<mgt-login id="login-two"></mgt-login>
<mgt-login id="login-three"></mgt-login>

<h2>Style using <code>applyTheme</code> function </h2>
<p>
  The first picker component is in the default light theme on a light
  background. We are setting the "picker-two" and "login-three" components to the dark theme using JavaScript.
  Because the login component is transparent it is necessary to set the background color using CSS.<br/>
  Refer to the CSS and JS tabs for details.
</p>
<mgt-people-picker id="picker-one"></mgt-people-picker>
<br>
<mgt-people-picker id="picker-two"></mgt-people-picker>
<style>
body {
  background-color: var(--fill-color);
  font-family: var(--body-font);
}
#login-two {
  background-color: aliceblue;
}
#login-three {
  background-color: var(--fill-color);
}
</style>

<script>
import { applyTheme } from '@microsoft/mgt-components';
applyTheme('light');
const darkElements = [
  document.querySelector("#picker-two"),
  document.querySelector("#login-three")
];
for (const element of darkElements) {
  if(element){
    // apply the dark theme on the second picker component only.
    applyTheme('dark', element)
  }
}
</script>
`;

export const localization = () => html`
<mgt-theme-toggle></mgt-theme-toggle>
  <style>
body {
    background-color: var(--fill-color);
}
  </style>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
  LocalizationHelper.strings = {
    _components: {
      "theme-toggle": {
        label: 'Theme 🎨:',
        on: 'Late night 🌙',
        off: 'Midday ☀️'
      },
    }
  }
  </script>
`;

export const events = () => html`
  <mgt-theme-toggle></mgt-theme-toggle>
  <style>
body {
    background-color: var(--fill-color);
}
  </style>
  <script>
    const themeToggle = document.querySelector('mgt-theme-toggle');
    themeToggle.addEventListener('updated', (e) => {
      console.log('updated', e);
    });
    themeToggle.addEventListener('darkmodechanged', (e) => {
      console.log('darkmodechanged', e);
    });
  </script>
`;
