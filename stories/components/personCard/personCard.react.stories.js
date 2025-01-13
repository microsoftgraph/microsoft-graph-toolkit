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
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { PersonCard } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      const onExpanded = useCallback((e: CustomEvent<null>) => {
        console.log("expanded");
      }, []);

      return (
        <PersonCard 
        personQuery="me" 
        updated={onUpdated}
        expanded={onExpanded}>
    </PersonCard>
      );
    };
  </react>
  <script>
    const personCard = document.querySelector('mgt-person-card');
    personCard.addEventListener('updated', (e) => {
      console.log("updated", e);
    });

    personCard.addEventListener('expanded', () => {
      console.log("expanded");
    })
  </script>
`;
