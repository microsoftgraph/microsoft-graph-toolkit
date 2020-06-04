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
import '../../dist/es6/components/mgt-people-picker/mgt-people-picker';

export default {
  title: 'Components | mgt-people-picker',
  component: 'mgt-people-picker',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const peoplePicker = () =>
  html`
    <mgt-people-picker></mgt-people-picker>
  `;

export const groupId = () => html`
  <mgt-people-picker group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people-picker>
`;

export const DarkMode = () => html`
  <mgt-people-picker></mgt-people-picker>
  <style>
    .story-mgt-preview-wrapper {
      background-color: black;
    }
    mgt-people-picker {
      --input-border: 2px rgba(255, 255, 255, 0.5) solid;
      --input-background-color: #1f1f1f;
      --dropdown-background-color: #1f1f1f;
      --dropdown-item-hover-background: #333d47;
      --dropdown-item-selected-background: #0f78d4;
      --input-hover-color: #008394;
      --input-focus-color: #0f78d4;
      --font-color: white;
      --placeholder-focus-color: rgba(255, 255, 255, 0.8);
    }
  </style>
`;

export const pickPeopleAndGroups = () => html`
  <mgt-people-picker type="any"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickGroups = () => html`
  <mgt-people-picker type="group"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickDistributionGroups = () => html`
  <mgt-people-picker type="group" group-type="distribution"></mgt-people-picker>
  <!-- group-type can be "any", "unified", "security", "mailenabledsecurity", "distribution" -->
`;
