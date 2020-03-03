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
import '../../dist/es6/components/mgt-person-card/mgt-person-card';
import '../../dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-person-card',
  component: 'mgt-person',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const personCard_hover = () => html`
  <mgt-person person-query="me" show-name show-email person-card="hover"></mgt-person>

  <div style="margin:2em 0 0 1em;font-family:segoe ui;color:#323130;font-size:12px">
    (Hover on person to view Person Card)
  </div>
`;

export const personCard_inheritDetails = () => html`
  <mgt-person person-query="me" show-name show-email person-card="hover">
    <template data-type="person-card">
      <mgt-person-card inherit-details></mgt-person-card>
    </template>
  </mgt-person>

  <div style="margin:2em 0 0 1em;font-family:segoe ui;color:#323130;font-size:12px">
    (Hover on person to view Person Card)
  </div>
`;
