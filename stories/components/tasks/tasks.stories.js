/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-tasks',
  component: 'tasks',
  decorators: [withCodeEditor],
  tags: ['autodocs'],
  parameters: {
    docs: {
      source: { code: '<mgt-tasks></mgt-tasks>' }
    }
  }
};

export const tasks = () => html`
  <mgt-tasks></mgt-tasks>
`;

export const tasksWithGroupId = () => html`
  <mgt-tasks group-id="45327068-6785-4073-8553-a750d6c16a45"></mgt-tasks>
  <!--
    NOTE: the default sandbox tenant doesn't have the required Tasks.ReadWrite and
    Group.ReadWrite.All permissions. Test this component in your tenant.
  -->
`;

export const useToDoAsDataSource = () => html`
  <mgt-tasks data-source="todo"></mgt-tasks>
`;
