/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / Templating',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const defaultTemplates = () => html`
  <mgt-person-card person-query="me">
    <template data-type="default">
      <div>
        <h3>user:</h3><p>{{this.person.givenName}}</p>
        <h3>mail:</h3><p>{{this.person.mail}}</p>
      </div>
    </template>
  </mgt-person-card>

`;

export const personDetails = () => html`
    <mgt-person person-query="me" view="twolines" person-card="hover">
      <template data-type="person-card">
        <mgt-person-card inherit-details>
          <template data-type="person-details">
          <div>
            <h3>user:</h3><p>{{this.person.givenName}}</p>
            <h3>mail:</h3><p>{{this.person.mail}}</p>
          </div>
        </mgt-person-card>
      </template>
    </mgt-person>
`;

export const additionalDetails = () => html`
    <mgt-person person-query="me" view="twolines" person-card="hover">
      <template data-type="person-card">
        <mgt-person-card inherit-details>
          <template data-type="additional-details">
            <h3>Stuffed Animal Friends:</h3>
            <ul>
              <li>Giraffe</li>
              <li>lion</li>
              <li>Rabbit</li>
            </ul>
          </template>
        </mgt-person-card>
      </template>
    </mgt-person>
`;
