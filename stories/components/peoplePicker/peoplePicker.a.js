/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-people-picker',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const People = () => html`
   <mgt-people-picker></mgt-people-picker>
 `;

export const RTL = () => html`
   <mgt-people-picker dir="RTL"></mgt-people-picker>
 `;

export const selectionChangedEvent = () => html`
  <mgt-people-picker></mgt-people-picker>
  <!-- Check the console tab for results -->
  <script>
  document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => {
    for (let i=0; i <e.detail.length;i++){
      const person = e.detail[i];
      console.log(person.displayName)
    }
});
  </script>
`;
