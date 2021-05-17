/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-person',
  component: 'mgt-person',
  decorators: [withCodeEditor]
};

export const person = () => html`
   <mgt-person person-query="me" view="twoLines"></mgt-person>
 `;

export const personLineClickEvents = () => html`
  <div style="margin-bottom: 10px">Click on each line</div>
  <div class="example">
    <mgt-person person-query="me" view="threelines"></mgt-person>
  </div>

  <div class="output">no line clicked</div>

  <script>
    const person = document.querySelector('mgt-person');
    person.addEventListener('line1clicked', e => {
      const output = document.querySelector('.output');

      if (e && e.detail && e.detail.displayName) {
        output.innerHTML = '<b>line1clicked:</b> ' + e.detail.displayName;
      }
    });
    person.addEventListener('line2clicked', e => {
      const output = document.querySelector('.output');

      if (e && e.detail && e.detail.mail) {
        output.innerHTML = '<b>line2clicked:</b> ' + e.detail.mail;
      }
    });
    person.addEventListener('line3clicked', e => {
      const output = document.querySelector('.output');

      if (e && e.detail && e.detail.jobTitle) {
        output.innerHTML = '<b>line3clicked:</b> ' + e.detail.jobTitle;
      }
    });
  </script>

  <style>
    .example {
      margin-bottom: 20px;
    }
  </style>
`;

export const personCard = () => html`
  <div class="example">
    <div style="margin-bottom:10px">Person card Hover</div>
    <mgt-person person-query="me" view="twoLines" person-card="hover"></mgt-person>
  </div>
  <div class="example">
  <div style="margin-bottom:10px">Person card Click</div>
    <mgt-person person-query="me" view="twoLines" person-card="click"></mgt-person>
  </div>
`;

export const RTL = () => html`
  <mgt-person person-query="me" view="twoLines" dir="RTL"></mgt-person>
`;

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

export const moreExamples = () => html`
  <style>
    .example {
      margin-bottom: 20px;
    }

    .styled-person {
      --font-family: 'Comic Sans MS', cursive, sans-serif;
      --color-sub1: red;
      --avatar-size: 60px;
      --font-size: 20px;
      --line2-color: green;
      --avatar-border-radius: 10% 35%;
      --line2-text-transform: uppercase;
    }

    .person-initials {
      --initials-color: yellow;
      --initials-background-color: red;
      --avatar-size: 60px;
      --avatar-border-radius: 10% 35%;
    }
  </style>

  <div class="example">
    <div>Default person</div>
    <mgt-person person-query="me"></mgt-person>
  </div>

  <div class="example">
    <div>One line</div>
    <mgt-person person-query="me" view="oneline"></mgt-person>
  </div>

  <div class="example">
    <div>Two lines</div>
    <mgt-person person-query="me" view="twoLines"></mgt-person>
  </div>

  <div class="example">
    <div>Change line content</div>
    <!--add fallback property by comma separating-->
    <mgt-person
      person-query="me"
      line1-property="givenName"
      line2-property="jobTitle,mail"
      view="twoLines"
    ></mgt-person>
  </div>

  <div class="example">
    <div>Large avatar</div>
    <mgt-person person-query="me" avatar-size="large"></mgt-person>
  </div>

  <div class="example">
    <div>Different styles (see css tab for style)</div>
    <mgt-person class="styled-person" person-query="me" view="twoLines"></mgt-person>
  </div>

  <div class="example" style="width: 200px">
    <div>Overflow</div>
    <mgt-person person-query="me" view="twoLines"></mgt-person>
  </div>

  <div class="example">
    <div>Style initials (see css tab for style)</div>
    <mgt-person class="person-initials" person-query="alex@fineartschool.net" view="oneline"></mgt-person>
  </div>

  <div>
    <div>Additional Person properties</div>
    <mgt-person
      person-query="me"
      view="twoLines"
      line1-property="displayName"
      line2-property="officeLocation"
    ></mgt-person>
  </div>
`;
