/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples | Accessibility',
  component: 'mgt-combo',
  decorators: [withSignIn, withCodeEditor],
  parameters: {
    a11y: {
      disabled: true
    },
    signInAddon: {
      test: 'test'
    }
  }
};

export const RightToLeft = () => html`
  <div class="dropdown">
    <label for="direction">Document Direction</label>
    <select id="direction" name="direction">
      <option value="ltr">Left-To-Right</option>
      <option value="rtl">Right-To-Left</option>
    </select>
  </div>
  <mgt-login></mgt-login>
  <mgt-person person-query="me"></mgt-person>
  <mgt-people-picker></mgt-people-picker>
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <mgt-tasks></mgt-tasks>
  <mgt-agenda></mgt-agenda>
  <mgt-people></mgt-people>
  <mgt-todo></mgt-todo>
  <script>
    let direction = document.querySelector('#direction');
    direction.addEventListener('change', function(e) {
      document.body.setAttribute('dir', e.target.value);
    });
  </script>
`;
