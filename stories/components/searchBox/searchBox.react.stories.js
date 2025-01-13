/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-box / React',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const searchBox = () => html`
  <mgt-search-box></mgt-search-box>
  <react>
    import { SearchBox } from '@microsoft/mgt-react';

    export default () => (
      <SearchBox></SearchBox>
    );
  </react>
`;

export const events = () => html`
  <!-- Open dev console and change the search box value -->
  <!-- See js tab for event subscription -->
  <mgt-search-box></mgt-search-box>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { SearchBox } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e) => {
        console.log('updated', e); 
      });

      const onSearchTermChanged = useCallback((e: CustomEvent<string>) => {
        console.log(e.detail);
      }, []);

      return (
        <SearchBox 
        updated={onUpdated}
        searchTermChanged={onSearchTermChanged}>
    </SearchBox>;
    };
  </react>
  <script>
    const searchBox = document.querySelector('mgt-search-box');
    searchBox.addEventListener('updated', (e) => {
      console.log('updated', e);
    });
    searchBox.addEventListener('searchTermChanged', (e) => {
      console.log(e.detail);
    });
  </script>
`;
