/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person / Templating',
  component: 'person',
  decorators: [withCodeEditor]
};

export const DefaultTemplates = () => html`
  <mgt-person person-query="me">
    <template>
      <div>
        Hello, my name is: {{person.displayName}}
      </div>
    </template>
    <template data-type="loading">
      Loading
    </template>
  </mgt-person>
`;

export const noDataTemplate = () => html`
  <div>
    <div>No data template</div>
    <mgt-person>
      <template data-type="no-data">
        <div>No person</div>
      </template>
    </mgt-person>
  </div>
  `;

export const retemplateMetadata = () => html`
  <mgt-person person-query="me" view="threelines">
    <template data-type="line1">
      <div>
        Hello, my name is: {{person.displayName}}
      </div>
    </template>
    <template data-type="line2">
      <div>
        {{person.jobTitle}}
      </div>
    </template>
    <template data-type="line3">
      <div>
        Loves MGT
      </div>
    </template>
  </mgt-person>

  <br/>

  <mgt-person person-query="me" view="fourlines">
    <template data-type="line1">
      <div>
        Hello, my name is: {{person.displayName}}
      </div>
    </template>
    <template data-type="line2">
      <div>
        Musician
      </div>
    </template>
    <template data-type="line3">
      <div>
        Calif records
      </div>
    </template>
    <template data-type="line4">
      <div>
        Nairobi
      </div>
    </template>
  </mgt-person>
`;

export const personCard = () => html`
    <mgt-person person-query="me" view="twolines" person-card="hover">
      <template data-type="person-card">
        <!-- <mgt-person-card inherit-details></mgt-person-card> -->
        My custom person card experience
      </template>
    </mgt-person>
`;
