/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / Properties',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const setCardDetails = () => html`
  <mgt-person-card class="my-person-card"></mgt-person-card>
  <script>
    const personCard = document.querySelector('.my-person-card');

    personCard.personDetails = {
      displayName: 'Megan Bowen',
      jobTitle: 'CEO',
      mail: 'megan@contoso.com',
      businessPhones: ['423-555-0120'],
      mobilePhone: '424-555-0130',
    };

    // set image
    personCard.personImage = '';
  </script>
`;

export const inheritDetails = () => html`
  <style>
    .note {
      margin: 2em 0 0 1em;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" view="twolines" person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details></mgt-person-card>
    </template>
  </mgt-person>
  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const setUserId = () => html`
  <mgt-person-card user-id="2804bc07-1e1f-4938-9085-ce6d756a32d2"></mgt-person-card>
`;

export const personCardConfig = () => html`
<mgt-person-card person-query="me" show-presence></mgt-person-card>
<script>
  import { MgtPersonCardConfig } from  '@microsoft/mgt-components';
  MgtPersonCardConfig.sections.files = false;
</script>`;
