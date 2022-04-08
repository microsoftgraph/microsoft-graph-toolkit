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

export const events = () => html`
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

export const RTL = () => html`
  <body dir="rtl">
    <mgt-person person-query="me" view="twoLines"></mgt-person>

    <!-- RTL with vertical layout -->
    <div class="row">
      <mgt-person person-query="me" vertical-layout id="online" view="oneline" person-card="hover"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" vertical-layout id="online2" view="twolines" person-card="hover"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" vertical-layout id="online3" view="threelines" class="example"></mgt-person>
    </div>
    <div class="row">
      <mgt-person person-query="me" vertical-layout id="online4" view="fourLines" class="example"></mgt-person>
    </div>
  </body>
`;

export const personVertical = () => html`

<div class="row">
  <mgt-person person-query="me" vertical-layout id="online" view="oneline" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" vertical-layout id="online2" view="twolines" person-card="hover"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" vertical-layout id="online3" view="threelines" class="example"></mgt-person>
</div>
<div class="row">
  <mgt-person person-query="me" vertical-layout id="online4" view="fourLines" class="example"></mgt-person>
</div>
`;
