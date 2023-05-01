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
    background-color: var(--neutral-fill-rest);
    color: var(--neutral-foreground-rest);
    font-family: var(--body-font);
}
</style>
`;

export const darkModeOn = () => html`
  <mgt-theme-toggle mode="dark"></mgt-theme-toggle>
  <style>
body {
    background-color: var(--neutral-fill-rest);
}
  </style>
`;

export const lightModeOn = () => html`
  <mgt-theme-toggle mode="light"></mgt-theme-toggle>
  <style>
body {
    background-color: var(--neutral-fill-rest);
}
  </style>
`;

export const themingWithoutToggle = () => html`
  <mgt-login id="login-one"></mgt-login>
  <mgt-login id="login-two"></mgt-login>
  <mgt-login id="login-three"></mgt-login>

  <!-- The login components are in the default light theme on a light
  background. We are setting the component with id "login-two" to the dark
  theme colors using JavaScript. -->

  <script>
    import { applyTheme } from '@microsoft/mgt';
    const loginTwo = document.querySelector("#login-two");

    if(loginTwo){
      // apply the dark theme on the second login component only.
      applyTheme('dark', loginTwo)
    }
  </script>
`;

export const localization = () => html`
<mgt-theme-toggle></mgt-theme-toggle>
  <style>
body {
    background-color: var(--neutral-fill-rest);
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
