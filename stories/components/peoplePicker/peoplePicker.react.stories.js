/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / React',
  component: 'people-picker',
  decorators: [withCodeEditor]
};

export const peoplePicker = () => html`
   <mgt-people-picker></mgt-people-picker>
  <react>
    import { PeoplePicker } from '@microsoft/mgt-react';

    export default () => (
      <PeoplePicker></PeoplePicker>
    );
  </react>
 `;

export const events = () => html`
  <mgt-people-picker></mgt-people-picker>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { PeoplePicker, IDynamicPerson } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      }, []);

      const onSelectionChanged = useCallback((e: CustomEvent<IDynamicPerson[]>) => {
        console.log(e.detail);
      }, []);

      return (
        <PeoplePicker 
        updated={onUpdated}
        selectionChanged={onSelectionChanged}>
    </PeoplePicker>

      );
    };
  </react>
  <script>
    document.querySelector('mgt-people-picker').addEventListener('updated', e => {
      console.log('updated', e);
    });
    document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => {
      console.log('selectionChanged', e);
    });
  </script>
`;
