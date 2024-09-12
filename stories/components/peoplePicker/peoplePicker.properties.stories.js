/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / Properties',
  component: 'people-picker',
  decorators: [withCodeEditor]
};

export const setPlaceholder = () => html`
<mgt-people-picker placeholder="Select people"></mgt-people-picker>
`;

export const showPresence = () => html`
<mgt-people-picker show-presence></mgt-people-picker>
`;

export const personCard = () => html`
  <mgt-people-picker person-card="hover"></mgt-people-picker>
`;

export const groupId = () => html`
<mgt-people-picker group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people-picker>
`;

export const singleSelectMode = () => html`
<mgt-people-picker selection-mode="single"></mgt-people-picker>

<h2>Render in a modal and clear on opening the modal</h2>
<button aria-label="open modal" id="modal">Open modal</button>

<div id="modal-content">
    <mgt-people-picker id="modal-picker" selection-mode="single"></mgt-people-picker>
    <button aria-label="close modal" id="close-modal">X</button>
</div>

<h2> Single select with default selected user Ids</h2>
<mgt-people-picker default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85" selection-mode="single"></mgt-people-picker>

<h2> Single select with default selected group Ids</h2>
<mgt-people-picker default-selected-group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9" selection-mode="single"></mgt-people-picker>

<style>
#modal-content {
  height: 200px;
  width: 100%;
  background-color: beige;
  display: none;
  align-items: center;
  justify-content: center;
  gap: 6px;
}
</style>

<script>
const modal = document.getElementById("modal")
const closeModal = document.getElementById("close-modal")
const modalContent = document.getElementById("modal-content")
const modalPicker = document.getElementById("modal-picker")
modal.addEventListener('click', () => {
  modalContent.style.display = "flex"
  modalPicker.selectedPeople = []
  const input = modalPicker.shadowRoot.querySelector('fluent-text-field').shadowRoot.querySelector('input');
  input.focus();
})

closeModal.addEventListener('click', () => {
  modalContent.style.display = "none";
  modal.focus();
})
</script>
`;

export const disableSuggestions = () => html`
<h1>Disable suggestions</h1>
<mgt-people-picker disable-suggestions></mgt-people-picker>
<h1>Disable suggestions with default selected user Ids</h1>
<mgt-people-picker
default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225" disable-suggestions>
</mgt-people-picker>


  
`;

export const dynamicGroupId = () => html`
<div class="groups">
  <label class="notes">Pick a group:
    <select id="groupChooser">
        <option></option>
    </select>
  </label>
</div>
<mgt-people-picker id="picker"></mgt-people-picker>


<style>
  .notes {
    font-size: 12px;
  }

  .groups {
    max-width: 200px;
    margin-bottom: 16px;
  }

  #groupChooser {
    position: inherit;
    width: 100%;
    max-height: 300px;
    overflow: scroll;
    padding-left: 3px;
  }
</style>
<script type="module">
import { Providers, ProviderState } from '@microsoft/mgt-element';

let picker = document.getElementById('picker');
let groupChooser = document.getElementById('groupChooser');


loadGroups();
Providers.onProviderUpdated(loadGroups);

function loadGroups() {
  let provider = Providers.globalProvider;
  if(provider && provider.state === ProviderState.SignedIn) {
    let client = provider.graph.client;

    client
      .api('/groups')
      .get()
      .then(groups => {
        for(let group of groups.value) {
          const id = group.id;
          let option = document.createElement('option');
          option.setAttribute("value", id);
          option.innerText = group.displayName;
          groupChooser.appendChild(option);
        }
      });
  }
}

function setGroupValue(selected) {
  picker.setAttribute('group-id', selected);
}

groupChooser.addEventListener('change', function(e) {
  const selection = e.target.value;
  if (selection !== -1) {
      setGroupValue(selection);
  }
});
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

export const pickMultipleGroupsShowMax = () => html`
  <mgt-people-picker type="group" group-type="unified,security,mailenabledsecurity" show-max="3"></mgt-people-picker>
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

export const pickerDisableImages = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225" disable-images>
  </mgt-people-picker>
`;

export const pickerDefaultSelectedUserIds = () => html`
  <mgt-people-picker
    default-selected-user-ids="e3d0513b-449e-4198-ba6f-bd97ae7cae85, 40079818-3808-4585-903b-02605f061225">
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
  <!-- Type any email address and press comma(,), semicolon(;), or enter keys to add it -->
  <script type="module">
    const peoplePicker = document.querySelector('mgt-people-picker');
    peoplePicker.selectedPeople = [{mail: "any@mail.com", displayName: "any@mail.com"}]
  </script>
`;

export const pickerUserIds = () => html`
  <mgt-people-picker
      user-ids="2804bc07-1e1f-4938-9085-ce6d756a32d2 ,e8a02cc7-df4d-4778-956d-784cc9506e5a,c8913c86-ceea-4d39-b1ea-f63a5b675166">
  </mgt-people-picker>
`;

export const pickerUserFilters = () => html`
  <mgt-people-picker
    user-filters="startsWith(displayName,'a')"
    user-type="user">
  </mgt-people-picker>
`;

export const pickerPeopleFilters = () => html`
  <mgt-people-picker
      people-filters="jobTitle eq 'Retail Manager'">
  </mgt-people-picker>

  <mgt-people-picker
    people-filters="personType/class eq 'Person' and personType/subclass eq 'OrganizationUser'">
  </mgt-people-picker>
`;

export const pickerGroupFilters = () => html`
  <mgt-people-picker
    group-filters="startsWith(displayName, 'a')"
    type="group">
  </mgt-people-picker>
`;

export const pickerGroupIds = () => html`
  <!-- This should show all the users in the groups of the group IDs -->
  <mgt-people-picker
    group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9,f2861ed7-abca-4556-bf0c-39ddc717ad81">
  </mgt-people-picker>
`;

export const pickerGroupIdsWithUserFilters = () => html`
  <!-- This should return an empty result -->
  <mgt-people-picker
    group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9,f2861ed7-abca-4556-bf0c-39ddc717ad81"
    user-filters="startswith(displayName, 's')">
  </mgt-people-picker>
  <br>
  <!-- This should return a result. Search for 'wil' should return Alex Wilber -->
  <mgt-people-picker
    group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9,f2861ed7-abca-4556-bf0c-39ddc717ad81"
    user-filters="startswith(displayName, 'a')">
  </mgt-people-picker>
`;

export const pickerGroupIdsWithTypeGroup = () => html`
  <!-- This should show the groups in the group-ids as groups -->
  <mgt-people-picker
    group-ids="94cb7dd0-cb3b-49e0-ad15-4efeb3c7d3e9,f2861ed7-abca-4556-bf0c-39ddc717ad81"
    type="group">
  </mgt-people-picker>
`;

export const pickerWithAriaLabel = () => html`
  <!-- This will set the aria-label attribute on the input element of the combo box -->
  <mgt-people-picker
    aria-label="Type to search for a user or group"
  >
  </mgt-people-picker>
`;

export const asyncDefaultSelectedUserIds = () => {
  return html`

  <mgt-people-picker id="async-picker"></mgt-people-picker>
  <script type="module">
  window.setTimeout(() => {
    const userIds = 'e3d0513b-449e-4198-ba6f-bd97ae7cae85,40079818-3808-4585-903b-02605f061225'.split(
      ','
    );
    const picker = document.getElementById('async-picker');
    picker.defaultSelectedUserIds = userIds;
    console.log('defaultSelectedUserIds set');
  }, 2000);
  </script>
`;
};
