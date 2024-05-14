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
  --agenda-background-color: gold;
  --agenda-event-box-shadow: 0px 2px 30px;
  --agenda-event-margin: 0px 10px 40px 10px;
  --agenda-event-padding: 8px 0px;
  --agenda-event-background-color: yellow;
  --agenda-event-border: dotted 2px white;

  --agenda-header-margin: 3px;
  --agenda-header-font-size: 20px;
  --agenda-header-color: maroon;

  --agenda-event-time-font-size: 20px;
  --agenda-event-time-color: maroon;

  --agenda-event-subject-font-size: 12px;
  --agenda-event-subject-color: maroon;

  --agenda-event-location-font-size: 20px;
  --agenda-event-location-color: maroon;
  --agenda-event-attendees-color: maroon;

  --people-overflow-font-color: maroon;
}
  </style>
`;
