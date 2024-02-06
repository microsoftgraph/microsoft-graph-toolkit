/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-picker',
  component: 'picker',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: {
        code: '<mgt-picker resource="me/todo/lists" scopes="tasks.read" placeholder="Select a task list" key-name="displayName"></mgt-picker>'
      },
      editor: { hidden: true }
    }
  }
};

export const picker = () => html`
  <mgt-picker resource="me/todo/lists" scopes="tasks.read" placeholder="Select a task list" key-name="displayName"></mgt-picker>
`;
