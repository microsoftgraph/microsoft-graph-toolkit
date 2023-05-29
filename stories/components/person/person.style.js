/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person / Style',
  component: 'person',
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
    .person {
      --person-background-color: #616161;
      --person-background-border-radius: 30%;

      --person-avatar-size: 40px;
      --person-avatar-border: 3px solid yellow;
      --person-avatar-border-radius: 54%;
      --person-initials-text-color: white;
      --person-initials-background-color: blue;

      --person-line1-font-size: 20px;
      --person-line1-font-weight: 600;
      --person-line1-text-color: red;
      --person-line1-text-transform: capitalize;
      --person-line1-text-line-height: 22px;

      --person-line2-font-size: 18px;
      --person-line2-font-weight: 500;
      --person-line2-text-color: orange;
      --person-line2-text-transform: full-width;
      --person-line2-text-line-height: 20px;

      --person-line3-font-size: 16px;
      --person-line3-font-weight: 400;
      --person-line3-text-color: blue;
      --person-line3-text-transform: uppercase;
      --person-line3-text-line-height: 18px;

      --person-line4-font-size: 14px;
      --person-line4-font-weight: 300;
      --person-line4-text-color: black;
      --person-line4-text-transform: lowercase;
      --person-line4-text-line-height: 16px;

      --person-details-spacing: 30px;
      --person-details-bottom-spacing: 20px;
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

  <mgt-person class="person" person-query="me" view="fourlines" id="online" show-presence></mgt-person>
  <br>
  <mgt-person class="person" person-query="me" view="fourlines" avatar-type="initials" id="dnd" show-presence vertical-layout></mgt-person>
`;
