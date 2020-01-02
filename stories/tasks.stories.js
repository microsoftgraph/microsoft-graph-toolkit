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
import '../dist/es6/components/mgt-tasks/mgt-tasks';
import '../dist/es6/mock/mgt-mock-provider';
import '../dist/es6/mock/MockProvider';

export default {
  title: 'mgt-tasks',
  component: 'mgt-tasks',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs, withSignIn],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const tasks = () => html`
  <mgt-tasks></mgt-tasks>
`;
