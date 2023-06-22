/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { defaultDocsPage } from '../../../.storybook/story-elements/defaultDocsPage';

export default {
  title: 'Components / mgt-theme-toggle',
  component: 'theme-toggle',
  decorators: [
    withCodeEditor({
      disableThemeToggle: true
    })
  ],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: { code: '<mgt-theme-toggle></mgt-theme-toggle>' }
    }
  }
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
  <h2>Style the background of a component</h2>
  <p>The login components are in the default light theme on a light
  background. We are setting the background of the component with id "login-two"
  to the <code>aquamarine</code> color using CSS. Check the CSS tab.</p>
  <mgt-login id="login-one"></mgt-login>
  <mgt-login id="login-two"></mgt-login>
  <mgt-login id="login-three"></mgt-login>


  <h2>Style using <code>applyTheme</code> function </h2>
  <p>The picker components are in the default light theme on a light
  background. We are setting the background of the component with id "picker-two"
  to the dark theme using JavaScript. Check the JS tab.</p>
  <mgt-people-picker id="picker-one"></mgt-people-picker>
  <br>
  <mgt-people-picker id="picker-two"></mgt-people-picker>
  <br>
  <mgt-people-picker id="picker-three"></mgt-people-picker>
  <br>

  <style>
    #login-two {
      background-color: aquamarine;
    }
  </style>

  <script>
    import { applyTheme } from '@microsoft/mgt';
    const pickerTwo = document.querySelector("#picker-two");

    if(pickerTwo){
      // apply the dark theme on the second picker component only.
      applyTheme('dark', pickerTwo)
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
  import { LocalizationHelper } from '@microsoft/mgt';
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
