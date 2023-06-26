/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-picker / Templating',
  component: 'picker',
  decorators: [withCodeEditor]
};

export const LoadingTemplate = () => html`
  <mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" key-name="displayName">
    <div>Loading template</div>
    <template data-type="loading">
      Loading
    </template>
  </mgt-picker>
`;

export const noDataTemplate = () => html`
  <div>
    <div>No data template</div>
    <mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" key-name="displayName">
      <template data-type="no-data">
        <div>No data</div>
      </template>
    </mgt-picker>
  </div>
  `;

export const errorTemplate = () => html`
  <div>
    <div>Error template</div>
    <mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" key-name="displayName">
      <template data-type="error">
        <div>Error</div>
      </template>
    </mgt-picker>
  </div>
  `;
