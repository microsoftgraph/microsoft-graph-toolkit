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
import { withSignIn } from '../.storybook/signInAddon';
import { withCodeEditor } from '../.storybook/codeAddon';
import '../dist/es6/components/mgt-people-picker/mgt-people-picker';

export default {
  title: 'mgt-people-picker',
  component: 'mgt-people-picker',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const peoplePicker = () => html`
  <mgt-people-picker></mgt-people-picker>
`;
