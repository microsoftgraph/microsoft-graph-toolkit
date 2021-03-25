/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-tasks',
  component: 'mgt-tasks',
  decorators: [withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const tasks = () => html`
  <mgt-tasks></mgt-tasks>
`;

export const darkTheme = () => html`
  <mgt-tasks class="mgt-dark"></mgt-tasks>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
