/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-picker / React',
  component: 'picker',
  decorators: [withCodeEditor]
};

export const picker = () => html`
  <mgt-picker resource="me/todo/lists" scopes="tasks.read, tasks.readwrite" placeholder="Select a task list" key-name="displayName"></mgt-picker>
  <react>
    import { Picker } from '@microsoft/mgt-react';

    export default () => (
      <Picker resource='me/todo/lists' scopes={['tasks.read']} placeholder="Select a task list" keyName="displayName"></Picker>
    );
  </react>
`;

export const events = () => html`
  <!-- Inspect to view log -->
  <mgt-picker resource="me/messages" scopes="mail.read" placeholder="Select a message" key-name="subject" max-pages="2"></mgt-picker>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { Picker } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      const onSelectionChanged = useCallback((e: CustomEvent<any>) => {
        console.log('selectedItem', e.detail);
      }, []);

      return (
        <Picker
        resource='me/messages'
        scopes={['mail.read']}
        placeholder="Select a message"
        keyName="subject"
        maxPages={2}
        updated={onUpdated}
        selectionChanged={onSelectionChanged}>
    </Picker>
      );
    };
  </react>
  <script>
    document.querySelector('mgt-picker').addEventListener('updated', e => {
      console.log('updated', e);
    });
    document.querySelector('mgt-picker').addEventListener('selectionChanged', e => {
      console.log('selectedItem:', e.detail);
    });
  </script>
`;
