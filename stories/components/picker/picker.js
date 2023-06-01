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
  title: 'Components / mgt-picker',
  component: 'picker',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: {
        code:
          '<mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" placeholder="Select a task list" key-name="displayName"></mgt-picker>'
      }
    }
  }
};

export const picker = () => html`
  <mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" placeholder="Select a task list" key-name="displayName"></mgt-picker>
`;

export const MaxPages = () => html`
  <mgt-picker resource="me/messages" scopes="mail.read" placeholder="Select a message" key-name="subject" max-pages="2"></mgt-picker>
`;

export const SelectedValue = () => html`
  <mgt-picker resource="/groups" selected-value="Activewear" scopes="group.read.all" key-name="displayName"></mgt-picker>
`;
export const CacheEnabled = () => html`
  <mgt-picker resource="/groups" placeholder="Select a group" scopes="group.read.all" key-name="displayName" cache-enabled="true" cache-invalidation-period="50000"></mgt-picker>
`;

export const events = () => html`
  <!-- Inspect to view log -->
  <mgt-picker resource="me/messages" scopes="mail.read" placeholder="Select a message" key-name="subject" max-pages="2"></mgt-picker>
  <script>
    document.querySelector('mgt-picker').addEventListener('selectionChanged', e => {
      console.log('selectedItem:', e.detail)
    });
  </script>
`;
