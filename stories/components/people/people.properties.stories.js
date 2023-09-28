/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / Properties',
  component: 'people',
  decorators: [withCodeEditor]
};

export const ShowMax = () => html`
  <mgt-people show-max="5"></mgt-people>
`;

export const ShowPresence = () => html`
  <mgt-people show-presence></mgt-people>
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

export const PeopleResource = () => html`
  <mgt-people resource="/me/directReports"></mgt-people>
`;

export const PersonCard = () => html`
  <div style="margin-bottom:10px">Person card Hover</div>
  <mgt-people show-max="5" person-card="hover"></mgt-people>
  <div style="margin-bottom:10px">Person card Click</div>
  <mgt-people show-max="5" person-card="click"></mgt-people>
`;

export const FallbackDetails = () => html`
  <mgt-people
    user-ids="2804bc07-1e1f-4938-9085-ce6d756a32d2, validButNotFound@email.in.ad"
    fallback-details='[{"mail": "test@test.com"},{"displayName": "validButNotFound@email.in.ad", "mail":"validButNotFound@email.in.ad"}]'
  >
  </mgt-people>
`;

export const FallbackDetailsPeopleQuery = () => html`
  <mgt-people
  people-queries="LidiaH@M365x214355.onmicrosoft.com, test@test.com"
  fallback-details='[{ "mail":"testmail@test.com"},{"mail": "test@test.com"}]'
  >
  </mgt-people>
`;
