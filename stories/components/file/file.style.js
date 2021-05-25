/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file / Style',
  component: 'mgt-file',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
    <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" class="mgt-dark"></mgt-file>
`;

export const customCSSProperties = () => html`
  <style>
    mgt-file {
    --file-type-icon-size: 30px;
    --file-border: 4px dotted #ffbdc3;
    --file-box-shadow: none;
    --file-background-color: #e0f8db;
    --font-family: 'Comic Sans MS', cursive, sans-serif;;
    --font-size: 15px;
    --font-weight: 2px;
    --text-transform: "";
    --color: black;
    --line2-font-size: 11px;
    --line2-font-weight	: 3px;
    --line2-color: red;
    --line2-text-transform: capitalize;	
    --line3-font-size: 12px;
    --line3-font-weight: 3px;
    --line3-color: purple;
    --line3-text-transform: lowercase;
    }
  </style>
  <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
`;
