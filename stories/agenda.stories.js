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
import '../dist/es6/components/mgt-agenda/mgt-agenda';

export default {
  title: 'mgt-agenda',
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

export const getByEventTemplate = () => html`
  <mgt-agenda show-max="7" days="10">
    <template data-type="event">
      <div class="root">
        <div class="time-container">
          <div class="date">{{{ dayFromDateTime(event.start.dateTime)}}}</div>
          <div class="time">{{{ timeRangeFromEvent(event) }}}</div>
        </div>

        <div class="separator">
          <div class="vertical-line top"></div>
          <div class="circle">
            <div data-if="!event.bodyPreview.includes('Join Microsoft Teams Meeting')" class="inner-circle"></div>
          </div>
          <div class="vertical-line bottom"></div>
        </div>

        <div class="details">
          <div class="subject">{{ event.subject }}</div>
          <div class="location" data-if="event.location.displayName">
            at
            <a href="https://bing.com/maps/default.aspx?where1={{event.location.displayName}}" target="_blank"
              ><b>{{ event.location.displayName }}</b></a
            >
          </div>
          <div class="attendees" data-if="event.attendees.length">
            <span class="attendee" data-for="attendee in event.attendees">
              <mgt-person person-query="{{attendee.emailAddress.name}}"></mgt-person>
            </span>
          </div>
          <div class="online-meeting" data-if="event.bodyPreview.includes('Join Microsoft Teams Meeting')">
            <img class="online-meeting-icon" src="https://img.icons8.com/color/48/000000/microsoft-teams.png" />
            <a class="online-meeting-link" href="{{ event.onlineMeetingUrl }}">Join Teams Meeting</a>
          </div>
        </div>
      </div>
    </template>
  </mgt-agenda>

  <script>
    let agenda = document.querySelector('mgt-agenda');

    agenda.templateConverters = agenda.templateConverters || {};
    agenda.templateConverters.dayFromDateTime = dateTimeString => {
      let date = new Date(dateTimeString);
      date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
      let monthNames = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
      ];

      let monthIndex = date.getMonth();
      let day = date.getDate();
      let year = date.getFullYear();

      return monthNames[monthIndex] + ' ' + day + ' ' + year;
    };

    agenda.templateConverters.timeRangeFromEvent = event => {
      if (event.isAllDay) {
        return 'ALL DAY';
      }

      let start = prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
      let end = prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

      return start + ' - ' + end;
    };

    function prettyPrintTimeFromDateTime(date) {
      date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
      let hours = date.getHours();
      let minutes = date.getMinutes();
      let ampm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12;
      hours = hours ? hours : 12;
      let minutesStr = minutes < 10 ? '0' + minutes : minutes;
      return hours + ':' + minutesStr + ' ' + ampm;
    }
  </script>

  <style>
    .root {
      display: flex;
      min-height: 75px;
      background: white;
    }

    .time-container {
      width: 200px;
      margin-top: 10px;
      margin-right: 10px;
      display: flex;
      flex-direction: column;
      align-items: flex-end;
    }

    .date {
      font-size: 13px;
      font-weight: bold;
    }

    .time {
      display: inline;
      font-size: 12px;
    }

    .location {
      font-size: 12px;
    }

    .details {
      margin-bottom: 20px;
      margin-top: 5px;
      font-size: 18px;
      min-width: 180px;
      padding-left: 5px;
      width: 400px;
    }

    .separator {
      width: 40px;
      display: flex;
      flex-direction: column;
      align-items: center;
      opacity: 1;
    }

    .vertical-line {
      align-content: center;
      width: 4px;
      background-color: #e5f2f3;
    }

    .vertical-line.top {
      height: 16px;
    }

    .vertical-line.bottom {
      flex: 1;
    }

    .circle {
      border-radius: 50%;
      height: 20px;
      width: 20px;
      position: relative;
      border: 2px solid #e5f2f3;
    }

    .circle .inner-circle {
      position: absolute;
      background-color: #307176;
      border-radius: 50%;
      height: 8px;
      width: 8px;

      top: 50%;
      left: 50%;
      margin: -4px 0 0 -4px;
    }

    .online-meeting {
      font-size: 10px;
      margin-top: 8px;
    }

    .online-meeting-icon {
      max-width: 16px;
      vertical-align: middle;
    }

    .online-meeting-link {
      text-decoration: none;
      color: #3063b2;
    }

    .attendees {
      margin-top: 8px;
    }

    .attendee {
      display: inline-block;
    }

    mgt-agenda {
      margin: 20px;
    }

    mgt-person {
      --avatar-size-s: 16px;
      margin-right: 4px;
    }
  </style>
`;
