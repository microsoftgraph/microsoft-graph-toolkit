/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / HTML',
  component: 'people-picker',
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

export const events = () => html`
   <mgt-people-picker></mgt-people-picker>
   <!-- Check the console tab for results -->
   <script>
   document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => {
     console.log(e.detail)
   });
   document.querySelector('mgt-people-picker').addEventListener('updated', e => {
      console.log('updated', e);
    });
   </script>
 `;

export const getEntityType = () => html`
  <mgt-people-picker
    type="person">
  </mgt-people-picker>
    <!-- Group entityType -->
    <!-- <mgt-people-picker type="group"></mgt-people-picker> -->

    <div class="entity-type"></div>

  <script type="module">
  import { isUser, isContact, isGroup } from '@microsoft/mgt-components';
  let entityType;
  const output = document.querySelector('.entity-type');
  const handleSelection = (e) => {
    const selected = e.detail[0];
    if (isGroup(selected)) {
        entityType = 'group'
    } else if (isUser(selected)) {
        entityType = 'user'
    } else if (isContact(selected)) {
        entityType = 'contact'
    }
    output.innerHTML = '<b>entityType:</b>' + entityType;
  }
  document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => handleSelection(e));

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
   import { LocalizationHelper } from '@microsoft/mgt-element';
   LocalizationHelper.strings = {
     _components: {
       'people-picker': {
          inputPlaceholderText: 'Search for 🤼',
          noResultsFound: '🤷‍♀️',
          loadingMessage: '🦔',
          selected: '👉',
          removeSelectedUser: '❌ ',
          selectContact: 'Select 🤼',
          suggestionsTitle: '📃 Suggested 🤼'
        }
     }
   }
   </script>
 `;

export const copyOrCutToPaste = () => html`
   <mgt-people-picker
    default-selected-user-ids="e8a02cc7-df4d-4778-956d-784cc9506e5a,eeMcKFN0P0aANVSXFM_xFQ==,48d31887-5fad-4d73-a9f5-3c356e68a038,e3d0513b-449e-4198-ba6f-bd97ae7cae85">
  </mgt-people-picker>
   <br/>
   <mgt-people-picker></mgt-people-picker>
 `;

export const pasteListOfUserIds = () => html`
<mgt-people-picker></mgt-people-picker>
<!-- You can paste emails or user IDs that are delimited with a comma(",") or semi-colon(";") -->
<!-- 48d31887-5fad-4d73-a9f5-3c356e68a038,24fcbca3-c3e2-48bf-9ffc-c7f81b81483d -->
<!-- MeganB@M365x214355.onmicrosoft.com;martin@musale.com;BrianJ@M365x214355.onmicrosoft.com-->
`;
