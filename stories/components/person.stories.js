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
import '../../dist/es6/components/mgt-person/mgt-person';

export default {
  title: 'Components | mgt-person',
  component: 'mgt-person',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const person = () => html`
  <mgt-person person-query="me" show-name show-email></mgt-person>
`;

export const personPhotoOnly = () => html`
  <mgt-person person-query="me"></mgt-person>
`;

export const personCardHover = () => html`
  <mgt-person person-query="me" person-card="hover"></mgt-person>
`;

export const personCardClick = () => html`
  <mgt-person person-query="me" person-card="click"></mgt-person>
`;
