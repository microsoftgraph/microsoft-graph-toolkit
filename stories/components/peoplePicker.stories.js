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
import '../../packages/mgt/dist/es6/components/mgt-people-picker/mgt-people-picker';

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

export const theme = () => html`
  <div class="mgt-light">
    <header class="mgt-dark">
      <p>I should be dark, regional class</p>
      <mgt-people-picker></mgt-people-picker>
      <div class="mgt-light">
        <p>I should be light, second level regional class</p>
        <mgt-people-picker></mgt-people-picker>
      </div>
    </header>
    <article>
      <p>I should be light, global class</p>
      <mgt-people-picker></mgt-people-picker>
    </article>
    <p>I am custom themed</p>
    <mgt-people-picker class="custom1"></mgt-people-picker>
    <p>I have both custom input background color and mgt-dark theme</p>
    <mgt-people-picker class="mgt-dark custom2"></mgt-people-picker>
    <p>I should be light, with unknown class mgt-foo</p>
    <mgt-people-picker class="mgt-foo"></mgt-people-picker>
  </div>
  <style>
    .custom1 {
      --input-border: 2px solid teal;
      --input-background-color: #33c2c2;
      --dropdown-background-color: #33c2c2;
      --dropdown-item-hover-background: #2a7d88;
      --input-hover-color: #b911b1;
      --input-focus-color: #441540;
      --font-color: white;
      --placeholder-focus-color: #441540;
      --selected-person-background-color: #441540;
    }

    .custom2 {
      --input-background-color: #e47c4d;
    }
  </style>
`;

export const pickPeopleAndGroups = () => html`
  <mgt-people-picker type="any"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickPeopleAndGroupsNested = () => html`
  <mgt-people-picker type="any" transitive-search="true"></mgt-people-picker>
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

export const pickerOverflowGradient = () => html`
  <mgt-people-picker
    default-selected-user-ids="e8a02cc7-df4d-4778-956d-784cc9506e5a,eeMcKFN0P0aANVSXFM_xFQ==,48d31887-5fad-4d73-a9f5-3c356e68a038,e3d0513b-449e-4198-ba6f-bd97ae7cae85"
  ></mgt-people-picker>
  <style>
    .story-mgt-preview-wrapper {
      width: 120px;
    }
  </style>
`;

export const pickerDefaultSelectedUserIds = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225"
  ></mgt-people-picker>
`;
