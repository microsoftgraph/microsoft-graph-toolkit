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
import { withSignIn } from '../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';
import '../dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'mgt-person',
  component: 'mgt-person',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const person = () => html`
  <mgt-person person-query="me" show-name show-email></mgt-person>
`;

export const personCard_hover = () => html`
  <mgt-person person-query="me" person-card="hover"></mgt-person>
`;

export const personCard_inheritDetails = () => html`
  <mgt-person person-query="me" show-name show-email person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details></mgt-person-card>
    </template>
  </mgt-person>
`;

export const personCard_additionalDetails = () => html`
  <mgt-person person-query="me" show-name show-email person-card="hover">
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
