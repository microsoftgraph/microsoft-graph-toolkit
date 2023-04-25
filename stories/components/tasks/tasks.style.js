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
  title: 'Components / mgt-tasks',
  component: 'tasks',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: { code: '<mgt-tasks></mgt-tasks>' }
    }
  }
};

export const CustomCSSProperties = () => html`
  <style>
    mgt-tasks {
        --tasks-header-padding: 28px 14px;
        --tasks-header-margin: 0 14px;
        --tasks-header-font-size: large;
        --tasks-header-font-weight: 800;
        --tasks-header-text-color: blue;
        --tasks-header-text-hover-color: gray;

        --tasks-new-button-width: 300px;
        --tasks-new-button-height: 50px;
        --tasks-new-button-text-color: orange;
        --tasks-new-button-text-font-weight: 400;
        --tasks-new-button-background: black;
        --tasks-new-button-border: 2px dotted golden;
        --tasks-new-button-background-hover: gray;
        --tasks-new-button-background-active: red;

        --task-complete-checkbox-background-color: red;
        --task-complete-checkbox-text-color: yellow;
        --task-incomplete-checkbox-background-color: orange;
        --task-incomplete-checkbox-background-hover-color: yellow;

        --task-title-text-font-size: large;
        --task-title-text-font-weight: 500;
        --task-complete-title-text-color: green;
        --task-incomplete-title-text-color: purple;

        --task-icons-width: 32px;
        --task-icons-height: 32px;
        --task-icons-background-color: purple;
        --task-icons-text-font-color: burlywood;
        --task-icons-text-font-size: 16px;
        --task-icons-text-font-weight: 400;

        --task-complete-background-color: powderblue;
        --task-incomplete-background-color: salmon;
        --task-complete-border: 2px dashed green;
        --task-incomplete-border: 2px double red;
        --task-complete-border-radius: 8px;
        --task-incomplete-border-radius: 12px;
        --task-complete-padding: 8px;
        --task-incomplete-padding: 12px;
        --tasks-gap: 8px;
    }
  </style>
  <mgt-tasks></mgt-tasks>
`;
