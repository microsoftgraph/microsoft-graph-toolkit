/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / React',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const personCard = () => html`
  <mgt-person-card person-query="me"></mgt-person-card>
<react>
import { PersonCard } from '@microsoft/mgt-react';

export default () => (
  <PersonCard personQuery="me"></PersonCard>
);
</react>
`;

export const events = () => html`
  <!-- Open dev console and click on an event -->
  <!-- See js tab for event subscription -->

  <mgt-person-card person-query="me"></mgt-person-card>
<react>
import { PersonCard } from '@microsoft/mgt-react';

const onExpanded = (e: CustomEvent<null>) => {
  console.log("expanded");
}

export default () => (
  <PersonCard personQuery="me" expanded={onExpanded}></PersonCard>
);
</react>
  <script>
    const personCard = document.querySelector('mgt-person-card');
    personCard.addEventListener('expanded', () => {
      console.log("expanded");
    })
  </script>
`;
