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

export const peoplePicker = () => html`
   <mgt-people-picker></mgt-people-picker>
 `;

export const RTL = () => html`
   <body dir="rtl">
     <mgt-people-picker></mgt-people-picker>
   </body>
 `;

export const selectionChangedEvent = () => html`
   <mgt-people-picker></mgt-people-picker>
   <!-- Check the console tab for results -->
   <script>
   document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => {
     console.log(e.detail)
   });
   </script>
 `;

export const selectGroupsById = () => html`
   <mgt-people-picker></mgt-people-picker>
   <!-- Check the js tab for example -->
   <script>
   document.querySelector('mgt-people-picker').selectGroupsById(["94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9", "f2861ed7-abca-4556-bf0c-39ddc717ad81"]);
   </script>
 `;

export const selectUsersById = () => html`
   <mgt-people-picker></mgt-people-picker>
   <!-- Check the js tab for example -->
   <script>
   document.querySelector('mgt-people-picker').selectUsersById(["e3d0513b-449e-4198-ba6f-bd97ae7cae85", "40079818-3808-4585-903b-02605f061225"]);
   </script>
 `;

export const localization = () => html`
   <mgt-people-picker></mgt-people-picker>
   <script>
   import { LocalizationHelper } from '@microsoft/mgt';
   LocalizationHelper.strings = {
     _components: {
       'people-picker': {
         inputPlaceholderText: 'Search for 🤼',
         noResultsFound: '🤷‍♀️',
         loadingMessage: '🦔'
       }
     }
   }
   </script>
 `;
