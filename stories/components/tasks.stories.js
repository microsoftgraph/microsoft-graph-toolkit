/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { versionInfo } from '../versionInfo';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  parameters: {
    version: versionInfo
  },
  title: 'Components / mgt-tasks',
  component: 'mgt-tasks',
  decorators: [withCodeEditor]
};

export const tasks = () => html`
  <mgt-tasks></mgt-tasks>
`;

export const tasksWithGroupId = () => html`
  <mgt-tasks group-id="45327068-6785-4073-8553-a750d6c16a45"></mgt-tasks>
`;

export const darkTheme = () => html`
  <mgt-tasks class="mgt-dark"></mgt-tasks>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
