/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-file',
  component: 'mgt-file',
  decorators: [withCodeEditor]
};

export const simple = () => html`
    <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2"></mgt-file>
 `;

export const events = () => html`
   <!-- Open dev console and click on an event -->
   <!-- See js tab for event subscription -->
 
   <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2">
    <template data-type="default">
      <div class="root">
       <span>Found the file {{file.name}}</span> 
      </div>
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
  </mgt-file>
   <script>
     const file = document.querySelector('mgt-file');
     file.addEventListener('templateRendered', (e) => {
       console.log(e);
     })
   </script>
 `;

export const RTL = () => html`
    <mgt-file file-query="/me/drive/items/01BYE5RZZFWGWWVNHHKVHYXE3OUJHGWCT2" dir="rtl"></mgt-file>
 `;
