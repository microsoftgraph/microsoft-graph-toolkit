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
import '../../dist/es6/components/mgt-people/mgt-people';

export default {
  title: 'Components | mgt-people',
  component: 'mgt-people',
  decorators: [withA11y, withSignIn, withCodeEditor],
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
  ></mgt-people>
`;

export const PeopleQuery = () => html`
  <mgt-people
    people-query="LidiaH@M365x214355.onmicrosoft.com, MeganB@M365x214355.onmicrosoft.com, LynneR@M365x214355.onmicrosoft.com, BrianJ@M365x214355.onmicrosoft.com, JoniS@M365x214355.onmicrosoft.com"
  ></mgt-people>
`;
