/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-todo',
  component: 'todo',
  decorators: [withCodeEditor],
  tags: ['autodocs'],
  parameters: {
    docs: {
      source: { code: '<mgt-todo></mgt-todo>' }
    }
  }
};

export const todos = () => html`
  <mgt-todo></mgt-todo>
`;

export const tasksWithTargetId = () => html`
  <mgt-todo target-id="AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAESAAA="></mgt-todo>
`;

export const tasksWithInitialId = () => html`
  <mgt-todo initial-id="AAMkAGVmMDEzMTM4LTZmYWUtNDdkNC1hMDZiLTU1OGY5OTZhYmY4OAAuAAAAAAAiQ8W967B7TKBjgx9rVEURAQAiIsqMbYjsT5e-T7KzowPTAAAAAAESAAA="></mgt-todo>
`;

export const ReadOnly = () => html`
  <mgt-todo read-only></mgt-todo>
`;
