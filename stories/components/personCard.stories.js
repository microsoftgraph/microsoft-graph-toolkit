/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt/dist/es6/components/mgt-person-card/mgt-person-card';
import '../../packages/mgt/dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-person-card',
  component: 'mgt-person',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const personCard = () => html`
  <mgt-person-card person-query="me"></mgt-person-card>
`;

export const personCardHover = () => html`
  <style>
    .note {
      margin: 2em 0 0 1em;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" view="twoLines" person-card="hover"></mgt-person>
  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const personCardInheritDetails = () => html`
  <style>
    .note {
      margin: 2em 0 0 1em;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" view="twoLines" person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details></mgt-person-card>
    </template>
  </mgt-person>

  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const personCardWithPresence = () => html`
<script>
  const available = {
    activity: 'Available',
    availability: 'Available',
    id: null
  }; 
  
  document.getElementById('with-presence').personPresence = available;
</script>
<style>
  .note {
    margin: 2em 0 0 1em;
    font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
      'Helvetica Neue', sans-serif;
    color: #323130;
    font-size: 12px;
  }
</style>
<mgt-person id="with-presence" person-query="me" person-card="hover" view="twoLines" show-presence></mgt-person>

<div class="note">
  (Hover on person to view Person Card)
</div>
`;
