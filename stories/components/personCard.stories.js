/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-person-card',
  component: 'mgt-person',
  decorators: [withCodeEditor]
};

export const personCard = () => html`
  <mgt-person-card person-query="me"></mgt-person-card>
`;

export const personCardHover = () => html`
  <style>
    .note {
      margin: 2em 0 0 1em;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" view="twoLines" person-card="hover"></mgt-person>
  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const personCardInheritDetails = () => html`
  <style>
    .note {
      margin: 2em 0 0 1em;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" view="twoLines" person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details></mgt-person-card>
    </template>
  </mgt-person>

  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const personCardWithPresence = () => html`
  <script>
    const available = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };

    document.getElementById('with-presence').personPresence = available;
  </script>
  <style>
    .note {
      margin: 2em 0 0 1em;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person id="with-presence" person-query="me" person-card="hover" view="twoLines" show-presence></mgt-person>

  <div class="note">
    (Hover on person to view Person Card)
  </div>
`;

export const darkTheme = () => html`
  <mgt-person-card person-query="me" class="mgt-dark"></mgt-person-card>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
export const ScopesAndConfigureSections = () => html`
  <script>
    import { MgtPersonCard } from '@microsoft/mgt';

    MgtPersonCard.config.useContactApis = false;

    MgtPersonCard.config.sections.mailMessages = true;
    MgtPersonCard.config.sections.files = true;
    MgtPersonCard.config.sections.profile = true;
    MgtPersonCard.config.sections.organization = true;

    // disable only "Works With" subsection under organization
    // MgtPersonCard.config.sections.organization = { showWorksWith: false };

    // change config above to see scopes update
    document.querySelector('.scopes').textContent = MgtPersonCard.getScopes();
  </script>
  <style>
    .note {
      margin: 2em;
      color: #323130;
      font-size: 12px;
    }
  </style>
  <mgt-person person-query="me" person-card="hover" view="twoLines" show-presence></mgt-person>

  <div class="note">
    (Hover on person to view Person Card)
  </div>

  <div>Scopes: <span class="scopes"></span></div>
`;
