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
