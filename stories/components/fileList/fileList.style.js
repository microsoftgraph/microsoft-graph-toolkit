/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / Style',
  component: 'mgt-file-list',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
    <mgt-file-list class="mgt-dark"></mgt-file-list>
  `;

export const customCSSProperties = () => html`
    <style>
      mgt-file-list {
        --file-list-background-color: #e0f8db;
        --file-item-background-color--hover: #caf1de;
        --file-item-background-color--active: #acddde;
        --file-list-border: 4px dotted #ffbdc3;
        --file-list-box-shadow: none;
        --file-list-padding: 0;
        --file-list-margin: 0;
        --file-item-border-radius: 12px;
        --file-item-margin: 2px 6px;
        --file-item-border-top: 4px dotted #ffbdc3;
        --file-item-border-left: 4px dotted #ffbdc3;
        --file-item-border-right: 4px dotted #ffbdc3;
        --file-item-border-bottom: 4px dotted #ffbdc3;
        --show-more-button-background-color: #fef8dd;
        --show-more-button-background-color--hover: #ffe7c7;
        --show-more-button-font-size: 14px;
        --show-more-button-padding: 16px;
        --show-more-button-border-bottom-right-radius: 12px;
        --show-more-button-border-bottom-left-radius: 12px;
      }
    </style>
    <mgt-file-list></mgt-file-list>
  `;
