/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda / HTML',
  component: 'agenda',
  decorators: [withCodeEditor]
};

export const agenda = () => html`
  <mgt-agenda></mgt-agenda>
`;

export const events = () => html`
  <!-- Open dev console and click on an event -->
  <!-- See js tab for event subscription -->

  <mgt-agenda></mgt-agenda>
  <script>
    const agenda = document.querySelector('mgt-agenda');
    agenda.addEventListener('updated', e => {
      console.log('updated', e);
    });

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
