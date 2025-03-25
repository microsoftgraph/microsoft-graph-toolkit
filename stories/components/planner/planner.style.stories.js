/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-planner / Style',
  component: 'planner',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <style>
    .tasks {
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

        --task-add-new-button-width: 60px;
        --task-add-new-button-height: 35px;
        --task-add-new-button-text-color: orange;
        --task-add-new-button-text-font-weight: 400;
        --task-add-new-button-background: black;
        --task-add-new-button-border: 2px dotted white;
        --task-add-new-button-background-hover: gray;
        --task-add-new-button-background-active: red;

        --task-cancel-new-button-width: 60px;
        --task-cancel-new-button-height: 35px;
        --task-cancel-new-button-text-color: yellow;
        --task-cancel-new-button-text-font-weight: 400;
        --task-cancel-new-button-background: red;
        --task-cancel-new-button-border: 2px dashed white;
        --task-cancel-new-button-background-hover: brown;
        --task-cancel-new-button-background-active: red;

        --task-complete-checkbox-background-color: red;
        --task-complete-checkbox-text-color: yellow;
        --task-incomplete-checkbox-background-color: orange;
        --task-incomplete-checkbox-background-hover-color: yellow;

        --task-title-text-font-size: large;
        --task-title-text-font-weight: 500;
        --task-complete-title-text-color: #066406;
        --task-incomplete-title-text-color: purple;

        --task-icons-width: 32px;
        --task-icons-height: 32px;
        --task-icons-background-color: purple;
        --task-icons-text-font-color: black;
        --task-icons-text-font-size: 16px;
        --task-icons-text-font-weight: 400;

        --task-complete-background-color: powderblue;
        --task-incomplete-background-color: salmon;
        --task-complete-border: 2px dashed #066406;
        --task-incomplete-border: 2px double red;
        --task-complete-border-radius: 8px;
        --task-incomplete-border-radius: 12px;
        --task-complete-padding: 8px;
        --task-incomplete-padding: 12px;
        --tasks-gap: 8px;

        --tasks-background-color: violet;
        --tasks-border: 2px dashed green;
        --tasks-border-radius: 12px;
        --tasks-padding: 16px;

        /** affects the date picker and the text-input field **/
        --task-new-input-border: 2px solid green;
        --task-new-input-border-radius: 8px;
        --task-new-input-background-color: yellow;
        --task-new-input-hover-background-color: yellowgreen;
        --task-new-input-placeholder-color: black;

        /** affects the date picker and the text-input field **/
        --task-new-dropdown-border: 2px solid green;
        --task-new-dropdown-border-radius: 8px;
        --task-new-dropdown-background-color: yellow;
        --task-new-dropdown-hover-background-color: yellowgreen;
        --task-new-dropdown-placeholder-color: red;
        --task-new-dropdown-option-text-color: blue;
        --task-new-dropdown-list-background-color: yellow;
        --task-new-dropdown-option-hover-background-color: yellowgreen;

        --task-new-person-icon-text-color: blue;
        --task-new-person-icon-color: blue;

        /** affects the options menu */
        --dot-options-menu-background-color: yellow;
        --dot-options-menu-shadow-color: yellow;
        --dot-options-menu-item-color: maroon;
        --dot-options-menu-item-hover-background-color: white;

        /** affects the assignments dropdown **/
        --arrow-options-button-font-color: #004074;
      }
  </style>
  <mgt-planner class="tasks"></mgt-planner>
`;
