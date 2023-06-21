/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples / General',
  decorators: [
    withCodeEditor({
      disableThemeToggle: true
    })
  ],
  parameters: {
    viewMode: 'story'
  }
};

export const theme = () => html`
<div>
  <p>This demonstrates how to set the theme globally without using a theme toggle and customize styling within specific scopes</p>
  <p>Please refer to the JS and CSS tabs in the editor for implentation details</p>
  <mgt-login></mgt-login>
  <p>This picker shows the custom focus ring color</p>
  <div class="custom-focus">
    <mgt-people-picker></mgt-people-picker>
  </div>
  <article>
    <p>I use the theme set on the body</p>
    <mgt-teams-channel-picker></mgt-teams-channel-picker>
  </article>
  <p>I am custom themed, take care to ensure that your customizations maintain accessibility standards</p>
  <mgt-teams-channel-picker class="custom1"></mgt-teams-channel-picker>
</div>
<script>
import { applyTheme } from '@microsoft/mgt';
const body = document.querySelector('body');
if(body) applyTheme('dark', body);
</script>
<style>
body {
  background-color: var(--fill-color);
  color: var(--neutral-foreground-rest);
  font-family: var(--body-font);
}
.custom-focus {
  --focus-ring-color: red;
  --focus-ring-style: solid;
}
.custom1 {
  --channel-picker-input-border: 2px solid teal;
  --channel-picker-input-background-color: black;
  --channel-picker-input-background-color-hover: #1a1a1a;
  --channel-picker-search-icon-color: yellow;
  --channel-picker-close-icon-color: yellow;
  --channel-picker-down-chevron-color: yellow;
  --channel-picker-up-chevron-color: yellow;
  --channel-picker-input-placeholder-text-color: white;
  --channel-picker-input-placeholder-text-color-focus: yellow;
  --channel-picker-input-placeholder-text-color-hover: chartreuse;
  --channel-picker-dropdown-background-color: #008383;
  --channel-picker-dropdown-item-background-color-hover: #006363;
  --channel-picker-font-color: white;
  --channel-picker-placeholder-default-color: white;
  --channel-picker-placeholder-focus-color: #441540;
}
</style>
`;
