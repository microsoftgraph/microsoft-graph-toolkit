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
import '../dist/es6/components/mgt-agenda/mgt-agenda';
import '../dist/es6/mock/mgt-mock-provider';
import '../dist/es6/mock/MockProvider';

export default {
  title: 'mgt-agenda',
  component: 'mgt-agenda',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const simple = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-agenda></mgt-agenda>
`;

export const getByEventQuery = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-agenda event-query="/me/events?orderby=start/dateTime"></mgt-agenda>
`;

export const getByDate = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-agenda group-by-day date="May 7, 2019" days="3"></mgt-agenda>
`;

export const getByEventTemplate = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-agenda>
    <template data-type="event">
      <button class="eventButton">
        <div class="event-subject">{{ event.subject }}</div>
        <div data-for="attendee in event.attendees">
          <mgt-person person-query="{{ attendee.emailAddress.name }}" show-name show-email> </mgt-person>
        </div>
      </button>
    </template>
    <template data-type="no-data">
      There are no events found!
    </template>
  </mgt-agenda>
`;
