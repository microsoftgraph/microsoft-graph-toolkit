/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file / Style',
  component: 'file',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <style>
    .file {
      /**NOTE: default-* tokens will override anything you set in the component.
      --default-font-size: 15px;
      --default-font-weight: 2px;
      */
      --default-font-family: 'Comic Sans MS', cursive, sans-serif;
      --file-type-icon-height:30px;
      --file-border: 4px dotted #ffbdc3;
      --file-border-radius: 8px;
      --file-box-shadow: 0px 2px 4px rgba(0, 0, 0, 0.1);
      --file-background-color: #e0f8db;
      --file-background-color-focus: yellow;
      --file-background-color-hover: green;
      --file-padding: 8px;
      --file-padding-inline-start: 12px;
      --file-margin: 3px 4px;
      --file-line1-font-size: 15px;
      --file-line1-font-weight: 500;
      --file-line1-color: blue;
      --file-line1-text-transform: capitalize;
      --file-line2-font-size:14px;
      --file-line2-font-weight:300;
      --file-line2-color: #e50000;
      --file-line2-text-transform: lowercase;
      --file-line3-font-size: 13px;
      --file-line3-font-weight: 500;
      --file-line3-color: purple;
      --file-line3-text-transform: capitalize;
    }
  </style>
  <mgt-file class="file" file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
`;
