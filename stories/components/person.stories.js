/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-person',
  component: 'mgt-person',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const person = () => html`
  <mgt-person person-query="me" show-name show-email></mgt-person>
`;

export const personPhotoOnly = () => html`
  <mgt-person person-query="me"></mgt-person>
`;

export const personPresence = () => html`
  <mgt-person person-query="me" show-presence show-email show-name></mgt-person>
`;

export const personPresenceDisplayAll = () => html`
  <script>
    const online = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };
    const onlineOof = {
      activity: 'OutOfOffice',
      availability: 'Available',
      id: null
    };
    const busy = {
      activity: 'Busy',
      availability: 'Busy',
      id: null
    };
    const busyOof = {
      activity: 'OutOfOffice',
      availability: 'Busy',
      id: null
    };
    const dnd = {
      activity: 'DoNotDisturb',
      availability: 'DoNotDisturb',
      id: null
    };
    const dndOof = {
      activity: 'OutOfOffice',
      availability: 'DoNotDisturb',
      id: null
    };
    const away = {
      activity: 'Away',
      availability: 'Away',
      id: null
    };
    const oof = {
      activity: 'OutOfOffice',
      availability: 'Offline',
      id: null
    };
    const offline = {
      activity: 'Offline',
      availability: 'Offline',
      id: null
    };

    const onlinePerson = document.getElementById('online');
    const onlineOofPerson = document.getElementById('onlineOof');
    const busyPerson = document.getElementById('busy');
    const busyOofPerson = document.getElementById('busyOof');
    const dndPerson = document.getElementById('dnd');
    const dndOofPerson = document.getElementById('dndOof');
    const awayPerson = document.getElementById('away');
    const oofPerson = document.getElementById('oof');
    const onlinePersonSmall = document.getElementById('online-small');
    const onlineOofPersonSmall = document.getElementById('onlineOof-small');
    const busyPersonSmall = document.getElementById('busy-small');
    const busyOofPersonSmall = document.getElementById('busyOof-small');
    const dndPersonSmall = document.getElementById('dnd-small');
    const dndOofPersonSmall = document.getElementById('dndOof-small');
    const awayPersonSmall = document.getElementById('away-small');
    const oofPersonSmall = document.getElementById('oof-small');

    onlinePerson.personPresence = online;
    onlineOofPerson.personPresence = onlineOof;
    busyPerson.personPresence = busy;
    busyOofPerson.personPresence = busyOof;
    dndPerson.personPresence = dnd;
    dndOofPerson.personPresence = dndOof;
    awayPerson.personPresence = away;
    oofPerson.personPresence = oof;
    onlinePersonSmall.personPresence = online;
    onlineOofPersonSmall.personPresence = onlineOof;
    busyPersonSmall.personPresence = busy;
    busyOofPersonSmall.personPresence = busyOof;
    dndPersonSmall.personPresence = dnd;
    dndOofPersonSmall.personPresence = dndOof;
    awayPersonSmall.personPresence = away;
    oofPersonSmall.personPresence = oof;
  </script>
  <style>
    mgt-person {
      display: block;
      margin: 12px;
    }
    mgt-person.small {
      display: inline-block;
      margin: 5px 0 0 10px;
    }
    .title {
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
      display: block;
      padding: 5px;
      font-size: 20px;
      margin: 10px 0 10px 0;
    }
    .title span {
      border-bottom: 1px solid #8a8886;
      padding-bottom: 5px;
    }
  </style>
  <div class="title"><span>Presence badge on big avatars: </span></div>
  <mgt-person id="online" person-query="me" show-presence show-email show-name></mgt-person>
  <mgt-person id="onlineOof" person-query="Isaiah Langer" show-presence show-email show-name></mgt-person>
  <mgt-person id="busy" person-query="bobk@tailspintoys.com" show-presence show-email show-name></mgt-person>
  <mgt-person id="busyOof" person-query="Diego Siciliani" show-presence show-email show-name></mgt-person>
  <mgt-person id="dnd" person-query="Lynne Robbins" show-presence show-email show-name></mgt-person>
  <mgt-person id="dndOof" person-query="EmilyB" show-presence show-email show-name></mgt-person>
  <mgt-person id="away" person-query="BrianJ" show-presence show-email show-name></mgt-person>
  <mgt-person id="oof" person-query="JoniS@M365x214355.onmicrosoft.com" show-presence show-email show-name></mgt-person>
  <div class="title"><span>Presence badge on small avatars: </span></div>
  <mgt-person class="small" id="online-small" person-query="me" show-presence></mgt-person>
  <mgt-person class="small" id="onlineOof-small" person-query="Isaiah Langer" show-presence></mgt-person>
  <mgt-person class="small" id="busy-small" person-query="bobk@tailspintoys.com" show-presence></mgt-person>
  <mgt-person class="small" id="busyOof-small" person-query="Diego Siciliani" show-presence></mgt-person>
  <mgt-person class="small" id="dnd-small" person-query="Lynne Robbins" show-presence></mgt-person>
  <mgt-person class="small" id="dndOof-small" person-query="EmilyB" show-presence></mgt-person>
  <mgt-person class="small" id="away-small" person-query="BrianJ" show-presence></mgt-person>
  <mgt-person class="small" id="oof-small" person-query="JoniS@M365x214355.onmicrosoft.com" show-presence></mgt-person>
`;

export const personCardHover = () => html`
  <mgt-person person-query="me" person-card="hover"></mgt-person>
`;

export const personCardClick = () => html`
  <mgt-person person-query="me" person-card="click"></mgt-person>
`;
