/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda / Style',
  component: 'agenda',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <mgt-agenda class="agenda" group-by-day></mgt-agenda>
  <style>
    .agenda {
      --agenda-background-color: #fcefef;
      --agenda-event-box-shadow: 0px 2px 30px pink;
      --agenda-event-margin: 0px 10px 40px 10px;
      --agenda-event-padding: 8px 0px;
      --agenda-event-background-color: #8d696f;
      --agenda-event-border: dotted 2px white;

      --agenda-header-margin: 3px;
      --agenda-header-font-size: 20px;
      --agenda-header-color: #7b575d;

      --agenda-event-time-font-size: 20px;
      --agenda-event-time-color: white;

      --agenda-event-subject-font-size: 12px;
      --agenda-event-subject-color: white;

      --agenda-event-location-font-size: 20px;
      --agenda-event-location-color: white;

      --agenda-event-attendees-color: gold;
    }
  </style>
`;
