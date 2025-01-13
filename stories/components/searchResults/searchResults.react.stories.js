/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results / React',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const searchResults = () => html`
  <mgt-search-results
    entity-types="driveItem"
    fetch-thumbnail="true"
    query-string="contoso">
  </mgt-search-results>
  <react>
    import { SearchResults } from '@microsoft/mgt-react';

    export default () => (
      <SearchResults entityTypes={['driveItem']} fetchThumbnail={true} queryString="contoso"></SearchResults>
    );
  </react>
`;

export const events = () => html`
  <mgt-search-results
    entity-types="driveItem"
    fetch-thumbnail="true"
    query-string="contoso">
  </mgt-search-results>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { SearchResults } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e) => {
        console.log('updated', e); 
      });

      const onDataChange = useCallback((e) => {
        console.log('dataChange', e); 
      });

      const onTemplateRendered = useCallback((e) => {
        console.log('templateRendered', e); 
      });

      return (
        <SearchResults 
        entityTypes={['driveItem']} 
        fetchThumbnail={true} 
        queryString="contoso"
        updated={onUpdated}
        dataChange={onDataChange}
        templateRendered={onTemplateRendered}>
      </SearchResults>
      );
    };
  </react>
  <script>
    const searchResults = document.querySelector('mgt-search-results');
    searchResults.addEventListener('updated', (e) => {
      console.log('updated', e);
    });
    searchResults.addEventListener('dataChange', (e) => {
      console.log('dataChange', e);
    });
    searchResults.addEventListener('templateRendered', (e) => {
      console.log('templateRendered', e);
    });
  </script>
`;
