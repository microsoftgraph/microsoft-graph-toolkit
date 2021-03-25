/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-people',
  component: 'mgt-people',
  decorators: [withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const People = () => html`
  <mgt-people show-max="5"></mgt-people>
`;

export const GroupId = () => html`
  <mgt-people group-id="02bd9fd6-8f93-4758-87c3-1fb73740a315"></mgt-people>
`;

export const UserIds = () => html`
  <mgt-people
    user-ids="2804bc07-1e1f-4938-9085-ce6d756a32d2 ,e8a02cc7-df4d-4778-956d-784cc9506e5a,c8913c86-ceea-4d39-b1ea-f63a5b675166"
  >
  </mgt-people>
`;

export const PeopleQueries = () => html`
  <mgt-people
    people-queries="LidiaH@M365x214355.onmicrosoft.com, Megan Bowen, Lynne Robbins, BrianJ@M365x214355.onmicrosoft.com, JoniS@M365x214355.onmicrosoft.com"
  >
  </mgt-people>
`;

export const darkTheme = () => html`
  <mgt-people class="mgt-dark"></mgt-people>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
