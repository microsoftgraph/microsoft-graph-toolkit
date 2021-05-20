/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / Style',
  component: 'mgt-person-card',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <mgt-person-card person-query="me" class="mgt-dark"></mgt-person-card>
  <style>
    body {
      background-color: black;
    }
  </style>
`;

export const customCSSProperties = () => html`
  <style>
    mgt-person-card {
      --person-card-display-name-font-size: 40px;
      --person-card-display-name-color: #ffffff;
      --person-card-title-font-size: 20px;
      --person-card-title-color: #ffffff;
      --person-card-subtitle-font-size: 10px;
      --person-card-subtitle-color: #ffffff;
      --person-card-details-title-font-size: 10px;
      --person-card-details-title-color: #b3bf0a;
      --person-card-details-item-font-size: 20px;
      --person-card-details-item-color: #3abf0a;
      --person-card-background-color: pink;
    }
  </style>
  <mgt-person-card person-query="me"></mgt-person-card>
`;
