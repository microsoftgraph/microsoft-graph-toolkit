/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-person / Style',
  component: 'mgt-person',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <div class="mgt-dark">
    <div class="title"><span>Transparent presence badge background:</span></div>
    <mgt-person person-query="me" view="twoLines" show-presence></mgt-person>
    <div class="title"><span>Light presence icon:</span></div>
    <mgt-person id="online" person-query="Isaiah Langer" show-presence view="twoLines"></mgt-person>
    <div class="title"><span>Dark presence icon:</span></div>
    <mgt-person id="dnd" person-query="Lynne Robbins" show-presence view="twoLines"></mgt-person>
  </div>
  <script>
    const online = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };
    const dnd = {
      activity: 'DoNotDisturb',
      availability: 'DoNotDisturb',
      id: null
    };
    const onlinePerson = document.getElementById('online');
    const dndPerson = document.getElementById('dnd');

    onlinePerson.personPresence = online;
    dndPerson.personPresence = dnd;
  </script>

  <style>
    body {
      background-color: black;
    }
    .title {
      color: white;
      display: block;
      padding: 5px;
      font-size: 20px;
      margin: 10px 0 10px 0;
      font-family: 'Segoe UI', 'Segoe UI Web (West European)', 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto,
        'Helvetica Neue', sans-serif;
    }
    .title span {
      border-bottom: 1px solid #8a8886;
      padding-bottom: 5px;
    }
    #online {
      --presence-icon-color: white;
    }
  </style>
`;

export const customCSSProperties = () => html`
  <style>
    mgt-person {
      --avatar-size: 90px;
      --avatar-border: 3px dotted red;
      --avatar-border-radius: 20% 40%;

      --initials-color: green;
      --initials-background-color: magenta;

      --presence-background-color: blue;
      --presence-icon-color: blue;

      --font-family: 'Segoe UI';
      --font-size: 25px;
      --font-weight: 700;
      --color-sub1: red;
      --text-transform: capitalize;

      --line2-font-size: 16px;
      --line2-font-weight: 400;
      --line2-color: green;
      --line2-text-transform: lowercase;

      --line3-font-size: 8px;
      --line3-font-weight: 400;
      --line3-color: pink;
      --line3-text-transform: none;

      --details-spacing: 30px;
    }
  </style>

  <script>
    const online = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };
    const dnd = {
      activity: 'DoNotDisturb',
      availability: 'DoNotDisturb',
      id: null
    };
    const onlinePerson = document.getElementById('online');
    const dndPerson = document.getElementById('dnd');

    onlinePerson.personPresence = online;
    dndPerson.personPresence = dnd;
  </script>

  <mgt-person person-query="me" view="threeLines" id="online" show-presence></mgt-person>
  <mgt-person person-query="me" view="threeLines" avatar-type="initials" id="dnd" show-presence></mgt-person>
`;
