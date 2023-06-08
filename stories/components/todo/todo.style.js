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
  title: 'Components / mgt-todo',
  component: 'todo',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: { code: '<mgt-todo></mgt-todo>' }
    }
  }
};

export const customCSSProperties = () => html`
  <style>
    .todo {
        --task-color: black;
        --task-background-color: white;
        --task-complete-background-color: grey;
        --task-date-input-active-color: blue;
        --task-date-input-hover-color: green;
        --task-background-color-hover: grey;
        --task-box-shadow: 0 0 10px 0 rgba(0, 0, 0, 0.2);
        --task-border: 1px solid black;
        --task-border-completed: 1px solid grey;
        --task-radio-background-color: green;
    }
  </style>
  <mgt-todo class="todo"></mgt-todo>
`;
