/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda',
  component: 'agenda',
  decorators: [withCodeEditor],
  tags: ['autodocs'],
  parameters: {
    docs: {
      source: { code: '<mgt-agenda></mgt-agenda>' }
    }
  }
};

export const simple = () => html`
  <mgt-agenda></mgt-agenda>
`;

export const events = () => html`
  <!-- Open dev console and click on an event -->
  <!-- See js tab for event subscription -->

  <mgt-agenda></mgt-agenda>
  <script>
    const agenda = document.querySelector('mgt-agenda');
    agenda.addEventListener('eventClick', (e) => {
      console.log(e.detail);
    })
  </script>
`;

export const RTL = () => html`
  <body dir="rtl">
   <mgt-agenda></mgt-agenda>
  </body>
`;
RTL.parameters = {
  docs: {
    source: {
      code: `
<body dir="rtl">
  <mgt-agenda></mgt-agenda>
</body>
`
    }
  }
};
