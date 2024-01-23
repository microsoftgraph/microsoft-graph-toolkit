/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples / Embed',
  decorators: [withCodeEditor],
  tags: ['hidden'],
  parameters: {
    viewMode: 'story'
  }
};

export const LoginToShowAgenda = () => html`
  <mgt-login></mgt-login>
  <mgt-agenda></mgt-agenda>
`;

export const LoginToShowAgendaReact = () => html`
  <mgt-login></mgt-login>
  <mgt-agenda></mgt-agenda>
  <react>
    import { Agenda, Login } from '@microsoft/mgt-react';

    export default () => (
      <>
        <Login></Login>
        <Agenda></Agenda>
      </>
    );
  </react>
`;
