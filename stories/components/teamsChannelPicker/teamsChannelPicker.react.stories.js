/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker / React',
  component: 'teams-channel-picker',
  decorators: [withCodeEditor]
};

export const teamsChannelPicker = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <react>
    import { TeamsChannelPicker } from '@microsoft/mgt-react';

    export default () => (
      <TeamsChannelPicker></TeamsChannelPicker>
    );
  </react>
`;

export const Events = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { TeamsChannelPicker, SelectedChannel } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('updated', e);
      });

      const onSelectionChanged = useCallback((e: CustomEvent<SelectedChannel | null>) => {
        console.log(e.detail);
      }, []);

      return (
        <TeamsChannelPicker 
        updated={onUpdated}
        selectionChanged={onSelectionChanged}>
    </TeamsChannelPicker>
      );
    };
  </react>
  <script>
    const picker = document.querySelector('mgt-teams-channel-picker');
    picker.addEventListener('selectionChanged', e => {
      console.log(e.detail);
    });
  </script>
`;
