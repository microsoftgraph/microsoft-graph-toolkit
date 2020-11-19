/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-teams-channel-picker/mgt-teams-channel-picker';

export default {
  title: 'Components | mgt-teams-channel-picker',
  component: 'mgt-teams-channel-picker',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const teamsChannelPicker = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
`;

export const getSelectedChannel = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <button class="button">Get SelectedChannel</button>

  <div class="output"></div>

  <script>
    document.querySelector('.button').addEventListener('click', _ => {
        const picker = document.querySelector('mgt-teams-channel-picker');
        const output = document.querySelector('.output');

        if (picker.selectedItem) {
          output.innerHTML = '<b>channel:</b> ' + picker.selectedItem.channel.displayName;
          output.innerHTML += '<br/><b>team:</b> ' + picker.selectedItem.team.displayName;
        } else {
          output.innerText = 'no channel selected';
        }
      });
  </script>
`;

export const selectionChangedEvent = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <div class="output">no channel selected</div>

  <script>
    const picker = document.querySelector('mgt-teams-channel-picker');
      picker.addEventListener('selectionChanged', e => {
        const output = document.querySelector('.output');

        if (e.detail.length) {
          output.innerHTML = '<b>channel:</b> ' + e.detail[0].channel.displayName;
          output.innerHTML += '<br/><b>team:</b> ' + e.detail[0].team.displayName;
        } else {
          output.innerText = 'no channel selected';
        }
      });
  </script>
`;

export const selectChannel = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>

  <button class="button">Select Channel By Id</button>

  <script>
    const button = document.querySelector('.button');
      button.addEventListener('click', async _ => {
        const picker = document.querySelector('mgt-teams-channel-picker');
        button.disabled = true;
        await picker.selectChannelById('19:d0bba23c2fc8413991125a43a54cc30e@thread.skype');
        button.disabled = false;
      });
  </script>
`;

export const theme = () => html`
  <div class="mgt-light">
    <header class="mgt-dark">
      <p>I should be dark, regional class</p>
      <mgt-teams-channel-picker></mgt-teams-channel-picker>
      <div class="mgt-light">
        <p>I should be light, second level regional class</p>
        <mgt-teams-channel-picker></mgt-teams-channel-picker>
      </div>
    </header>
    <article>
      <p>I should be light, global class</p>
      <mgt-teams-channel-picker></mgt-teams-channel-picker>
    </article>
    <p>I am custom themed</p>
    <mgt-teams-channel-picker class="custom1"></mgt-teams-channel-picker>
    <p>I have both custom input background color and mgt-dark theme</p>
    <mgt-teams-channel-picker class="mgt-dark custom2"></mgt-teams-channel-picker>
    <p>I should be light, with unknown class mgt-foo</p>
    <mgt-teams-channel-picker class="mgt-foo"></mgt-teams-channel-picker>
  </div>
  <style>
    .custom1 {
      --input-border: 2px solid teal;
      --input-background-color: #33c2c2;
      --dropdown-background-color: #33c2c2;
      --dropdown-item-hover-background: #2a7d88;
      --input-hover-color: #b911b1;
      --input-focus-color: #441540;
      --font-color: white;
      --placeholder-default-color: white;
      --placeholder-focus-color: #441540;
    }

    .custom2 {
      --input-background-color: #e47c4d;
    }
  </style>
`;
