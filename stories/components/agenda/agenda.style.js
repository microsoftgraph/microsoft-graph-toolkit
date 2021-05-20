/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda / Style',
  component: 'mgt-agenda',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <mgt-agenda class="mgt-dark"></mgt-agenda>
  <style>
    body {
      background-color: black;
    }
  </style>
`;

export const customProperties = () => html`
  <mgt-agenda group-by-day></mgt-agenda>
  <style>
    mgt-agenda {
      --event-box-shadow: 0px 2px 30px pink;
      --event-margin: 0px 10px 40px 10px;
      --event-padding: 8px 0px;
      --event-background-color: pink;
      --event-border: dotted 2px white;

      --agenda-header-margin: 3px;
      --agenda-header-font-size: 20px;
      --agenda-header-color: pink;

      --event-time-font-size: 20px;
      --event-time-color: white;

      --event-subject-font-size: 12px;
      --event-subject-color: white;

      --event-location-font-size: 20px;
      --event-location-color: white;
    }
  </style>
`;
