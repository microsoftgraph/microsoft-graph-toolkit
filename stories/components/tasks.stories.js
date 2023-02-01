/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import { versionInfo } from '../versionInfo';

export default {
  parameters: {
    version: versionInfo
  },
  title: 'Components / mgt-tasks',
  component: 'tasks',
  decorators: [withCodeEditor]
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
