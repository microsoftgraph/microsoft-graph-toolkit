/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / Properties',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const groupId = () => html`
  <mgt-people-picker group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people-picker>
`;

export const dynamicGroupId = () => html`
  <mgt-people-picker id="picker"></mgt-people-picker>
  <div>
    <p class="notes">Pick a group:</p>
    <select id="groupChooser" name="group">
      <option value="">none</option>
    </select>
    <p class="notes">People chosen:</p>
    <div id="chosenPeople"></div>
  </div>

  <style>
    body {
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
    }

    .notes {
      font-size: 12px;
      margin-bottom: 2px;
    }
  </style>
  <script type="module">
    import { Providers, ProviderState } from '@microsoft/mgt';

    let picker = document.getElementById('picker');
    let chosenArea = document.getElementById('chosenPeople');
    let groupChooser = document.getElementById('groupChooser');

    groupChooser.addEventListener('change', getGroupValue);

    loadGroups();
    Providers.onProviderUpdated(loadGroups);

    function loadGroups() {
      let provider = Providers.globalProvider;
      if (provider && provider.state === ProviderState.SignedIn) {
        let client = provider.graph.client;

        client
          .api('/groups')
          .get()
          .then(groups => {
            for (let group of groups.value) {
              let option = document.createElement('option');
              option.value = group.id;
              option.text = group.displayName;

              groupChooser.appendChild(option);
            }
          });
      }
    }

    picker.addEventListener('selectionChanged', function(e) {
      //reset area
      chosenArea.innerHTML = '';
      //render selected people to chosen people div
      for (var i = 0; i < e.detail.length; i++) {
        let newElem = document.createElement('div');
        newElem.innerHTML = e.detail[i].displayName + ' ' + e.detail[i].id;
        chosenArea.append(newElem);
      }
    });

    function getGroupValue(e) {
      let selection = groupChooser.selectedOptions[0];
      if (selection !== undefined) {
        setGroupValue(selection.value);
      }
    }

    function setGroupValue(selected) {
      picker.setAttribute('group', selected);
    }
  </script>
`;

export const pickPeopleAndGroups = () => html`
  <mgt-people-picker type="any"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickPeopleAndGroupsNested = () => html`
  <mgt-people-picker type="any" transitive-search="true"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickGroups = () => html`
  <mgt-people-picker type="group"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickDistributionGroups = () => html`
  <mgt-people-picker type="group" group-type="distribution"></mgt-people-picker>
  <!-- group-type can be "any", "unified", "security", "mailenabledsecurity", "distribution" -->
`;

export const pickMultipleGroups = () => html`
  <mgt-people-picker type="group" group-type="unified,distribution,security"></mgt-people-picker>
  <!-- group-type can be "any", "unified", "security", "mailenabledsecurity", "distribution" -->
`;

export const pickPeople = () => html`
  <mgt-people-picker type="person"></mgt-people-picker>
  <!-- type can be "any", "person", "group" -->
`;

export const pickOnlyOrganizationUsers = () => html`
  <mgt-people-picker type="person" user-type="user"></mgt-people-picker>
  <!-- user-type can be "any", "user", "contact" -->
`;

export const pickOnlyContacts = () => html`
  <mgt-people-picker type="person" user-type="contact"></mgt-people-picker>
  <!-- user-type can be "any", "user", "contact" -->
`;

export const pickerOverflowGradient = () => html`
  <mgt-people-picker
    default-selected-user-ids="e8a02cc7-df4d-4778-956d-784cc9506e5a,eeMcKFN0P0aANVSXFM_xFQ==,48d31887-5fad-4d73-a9f5-3c356e68a038,e3d0513b-449e-4198-ba6f-bd97ae7cae85"
  ></mgt-people-picker>
  <style>
    .story-mgt-preview-wrapper {
      width: 120px;
    }
  </style>
`;

export const pickerDisabled = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225" disabled>
  </mgt-people-picker>
`;

export const pickerDefaultSelectedUserIds = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225">
  </mgt-people-picker>
  </mgt-people-picker>
`;

export const pickerDefaultSelectedGroupIds = () => html`
  <mgt-people-picker
    default-selected-group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9, f2861ed7-abca-4556-bf0c-39ddc717ad81">
  </mgt-people-picker>
`;

export const pickerDefaultSelectedUserAndGroupIds = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225"
    default-selected-group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9, f2861ed7-abca-4556-bf0c-39ddc717ad81">
  </mgt-people-picker>
`;

export const pickerAllowAnyEmail = () => html`
  <mgt-people-picker allow-any-email></mgt-people-picker>
  <!-- Type any email address and press comma(,), semicolon(;), tab or enter to add it -->
  <script type="module">
    const peoplePicker = document.querySelector('mgt-people-picker');
    peoplePicker.selectedPeople = [{mail: "any@mail.com", displayName: "any@mail.com"}]
  </script>
`;
