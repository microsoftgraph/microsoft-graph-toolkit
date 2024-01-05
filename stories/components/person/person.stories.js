/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person',
  component: 'person',
  decorators: [withCodeEditor],
  tags: ['autodocs'],
  parameters: {
    docs: {
      source: { code: '<mgt-person person-query="me" view="twoLines"></mgt-person>' }
    }
  }
};

export const person = () => html`
<mgt-person person-query="me"></mgt-person>
<br>
<mgt-person person-query="me" view="oneLine"></mgt-person>
<br>
<mgt-person person-query="me" view="twoLines"></mgt-person>
<br>
<mgt-person person-query="me" view="threeLines"></mgt-person>
<br>
<mgt-person person-query="me" view="fourLines"></mgt-person>
`;

export const events = () => html`
  <div style="margin-bottom: 10px">Click on each line</div>
  <div class="example">
    <mgt-person person-query="me" view="fourlines"></mgt-person>
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

      if (e && e.detail && e.detail.jobTitle) {
          output.innerHTML = '<b>line2clicked:</b> ' + e.detail.jobTitle;
      }
  });
  person.addEventListener('line3clicked', e => {
      const output = document.querySelector('.output');

      if (e && e.detail && e.detail.department) {
          output.innerHTML = '<b>line3clicked:</b> ' + e.detail.department;
      }
  });
  person.addEventListener('line4clicked', e => {
      const output = document.querySelector('.output');

      if (e && e.detail && e.detail.mail) {
          output.innerHTML = '<b>line4clicked:</b> ' + e.detail.mail;
      }
  });

  </script>

  <style>
    .example {
      margin-bottom: 20px;
    }
  </style>
`;

export const RTL = () => html`
  <body dir="rtl">
    <div class="example">
      <mgt-person person-query="me" view="oneline"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="twolines"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="threelines"></mgt-person>
    </div>
    <div class="example">
      <mgt-person person-query="me" view="fourlines"></mgt-person>
    </div>

    <!-- RTL with vertical layout -->
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online" view="oneline"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online2" view="twolines"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online3" view="threelines"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" class="example" vertical-layout id="online4" view="fourLines"></mgt-person>
    </div>
  </body>
  <style>
  .example {
      margin-bottom: 20px;
    }
    </style>
`;

export const personVertical = () => html`

<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout view="oneline" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout view="twolines" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout view="threelines" class="example"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout view="fourLines" class="example"></mgt-person>
</div>

<!-- With Presence; Check JS tab -->

<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout id="online" show-presence view="oneline" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout id="online2" show-presence view="twolines" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout id="online3" show-presence view="threelines" class="example"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" class="example" vertical-layout id="online4" show-presence view="fourLines" class="example"></mgt-person>
</div>

<!-- Person unauthenticated vertical layout-->
<div class="row">
	<mgt-person person-query="mbowen" vertical-layout view="twoLines" fallback-details='{"mail":"MeganB@M365x214355.onmicrosoft.com"}'>
	</mgt-person>
</div>

<script>
            const online = {
          activity: 'Available',
          availability: 'Available',
          id: null
      };
      const onlinePerson = document.getElementById('online');
      const onlinePerson2 = document.getElementById('online2');
      const onlinePerson3 = document.getElementById('online3');
      const onlinePerson4 = document.getElementById('online4');

      onlinePerson.personPresence = online;
      onlinePerson2.personPresence = online;
      onlinePerson3.personPresence = online;
      onlinePerson4.personPresence = online;
    </script>
<style>
  .example {
      margin-bottom: 20px;
    }
    </style>
`;
