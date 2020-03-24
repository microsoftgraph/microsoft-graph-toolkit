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
import '../../dist/es6/components/mgt-agenda/mgt-agenda';

export default {
  title: 'Components | mgt-agenda',
  component: 'mgt-agenda',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const simple = () => html`
  <mgt-agenda></mgt-agenda>
`;

export const getByEventQuery = () => html`
  <mgt-agenda event-query="/me/events?orderby=start/dateTime"></mgt-agenda>
`;

export const getByDate = () => html`
  <mgt-agenda group-by-day date="May 7, 2019" days="3"></mgt-agenda>
`;

export const getEventsForNextWeek = () => html`
  <mgt-agenda group-by-day days="7"></mgt-agenda>
`;
