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
import '../dist/es6/components/mgt-people/mgt-people';
import '../dist/es6/mock/mgt-mock-provider';
import '../dist/es6/mock/MockProvider';

export default {
  title: 'mgt-people',
  component: 'mgt-people',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const People = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-people show-max="5"></mgt-people>
`;
