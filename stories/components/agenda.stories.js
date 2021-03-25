/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-agenda',
  component: 'mgt-agenda',
  decorators: [withCodeEditor],
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

export const darkTheme = () => html`
  <mgt-agenda group-by-day days="7" class="mgt-dark"></mgt-agenda>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
