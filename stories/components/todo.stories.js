/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-todo',
  component: 'mgt-todo',
  decorators: [withCodeEditor]
};

export const tasks = () => html`
  <mgt-todo></mgt-todo>
`;

export const ReadOnly = () => html`
  <mgt-todo read-only></mgt-todo>
`;

export const darkTheme = () => html`
  <mgt-todo class="mgt-dark"></mgt-todo>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
