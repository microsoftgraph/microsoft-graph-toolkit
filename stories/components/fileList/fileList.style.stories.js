/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / Style',
  component: 'file-list',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
    <style>
      .file-list {
        /** mgt-file-upload custom styling */
        --file-upload-background-color-drag: rgb(255, 0, 0, 0.5);
        --file-upload-border-drag: 2px groove black;
        --file-upload-button-background-color: purple;
        --file-upload-button-background-color-hover: green;
        --file-upload-button-text-color: whitesmoke;
        --file-upload-dialog-background-color: azure;
        --file-upload-dialog-text-color: yellow;
        --file-upload-dialog-replace-button-background-color: white;
        --file-upload-dialog-replace-button-background-color-hover: gray;
        --file-upload-dialog-replace-button-text-color: black;
        --file-upload-dialog-keep-both-button-background-color: black;
        --file-upload-dialog-keep-both-button-background-color-hover: gray;
        --file-upload-dialog-keep-both-button-text-color: white;
        --file-upload-button-border: 2px dotted yellow;
        --file-upload-dialog-replace-button-border: 2px dotted;
        --file-upload-dialog-keep-both-button-border: 2px dashed;
        --file-upload-dialog-border: 2px solid blue;
        --file-upload-dialog-width: 300px;
        --file-upload-dialog-height: 100px;
        --file-upload-dialog-padding: 36px;

        /** mgt-file custom styling */
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
        --file-line1-color: #727272;
        --file-line1-text-transform: capitalize;
        --file-line2-font-size:14px;
        --file-line2-font-weight:300;
        --file-line2-color: #e50000;
        --file-line2-text-transform: lowercase;
        --file-line3-font-size: 13px;
        --file-line3-font-weight: 500;
        --file-line3-color: purple;
        --file-line3-text-transform: capitalize;

        /** mgt-file-list CSS tokens */
        --file-list-background-color: #e0f8db;
        --file-list-box-shadow: none;
        --file-list-border: 4px dotted #ffbdc3;
        --file-list-border-radius: 10px;
        --file-list-padding: 0;
        --file-list-margin: 0;
        --show-more-button-background-color: #fef8dd;
        --show-more-button-background-color--hover: #ffe7c7;
        --show-more-button-font-size: 14px;
        --show-more-button-padding: 16px;
        --show-more-button-border-bottom-right-radius: 12px;
        --show-more-button-border-bottom-left-radius: 12px;
      }
    </style>
    <mgt-file-list class="file-list" enable-file-upload></mgt-file-list>
  `;
