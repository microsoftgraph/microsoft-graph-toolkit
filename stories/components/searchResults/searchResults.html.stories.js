/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results / HTML',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const searchResults = () => html`
  <mgt-search-results
    entity-types="driveItem"
    fetch-thumbnail="true"
    query-string="contoso">
  </mgt-search-results>
`;

export const localization = () => html`
  <mgt-search-results entity-types="driveItem" query-string="contoso">
  </mgt-search-results>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
  LocalizationHelper.strings = {
    _components: {
      'search-results': {
        modified: 'edited on',
        back: 'Previous',
        next: 'Next one',
        pages: 'sheets',
        page: 'Sheet'
      },
    }
  }
  </script>
`;

export const events = () => html`
  <mgt-search-results
    entity-types="driveItem"
    fetch-thumbnail="true"
    query-string="contoso">
  </mgt-search-results>
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
