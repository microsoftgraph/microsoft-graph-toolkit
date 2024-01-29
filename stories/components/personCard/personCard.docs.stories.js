/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card',
  component: 'person-card',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-person-card person-query="me" id="online" show-presence></mgt-person-card>' },
      editor: { hidden: true }
    }
  }
};

export const personCard = () => html`
  <mgt-person-card person-query="me" id="online" show-presence></mgt-person-card>

  <!-- Person Card without Presence -->
  <!-- <mgt-person-card person-query="me"></mgt-person-card> -->
  <script>
    const online = {
      activity: 'Available',
      availability: 'Available',
      id: null
    };
    const onlinePerson = document.getElementById('online');
    onlinePerson.personPresence = online;
  </script>
`;
