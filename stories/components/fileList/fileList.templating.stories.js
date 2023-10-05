/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list / Templating',
  component: 'file',
  decorators: [withCodeEditor]
};

export const defaultTemplates = () => html`
  <mgt-file-list>
    <template data-type="default">
      <div class="root" data-for="file in files">
        <span>Found the file {{file.name}}</span>
      </div>
        <p>The templateRendered event has been fired!</p>
    </template>
    <template data-type="loading">
      <div class="root">
        loading file
      </div>
    </template>
    <template data-type="no-data">
      <div class="root">
        there is no data
      </div>
    </template>
  </mgt-file-list>
`;

export const fileTemplate = () => html`
  <!-- Uncomment the CSS tab to see custom css on hover -->
  <mgt-file-list>
    <template data-type="file">
      <div>
        <span>Found the file {{file.name}}</span>
        <mgt-file data-props="fileDetails: file"></mgt-file>
      </div>
    </template>
  </mgt-file-list>
  <style>
    /* mgt-file-list {
      --file-item-background-color--hover: #caf1de;
      --file-item-background-color--active: #acddde;
    } */
  </style>
`;

export const fileTemplateEvents = () => html`
  <!-- You will see the selected file in the JS console -->
  <mgt-file-list id="fileList">
    <template data-type="file">
      <div>
        <span>Found the file {{file.name}}</span>
      </div>
    </template>
  </mgt-file-list>
  <script>
    document.getElementById('fileList').addEventListener('itemClick', e => {
      console.log('Item Clicked:', e.detail.name)
    })
  </script>
`;
