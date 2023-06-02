/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Preview / mgt-search-box',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const searchBox = () => html`
  <mgt-search-box>
  </mgt-search-box>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const localization = () => html`
  <mgt-search-box>
  </mgt-search-box>
  <script>
  import '@microsoft/mgt-components/dist/es6/components/preview';
  import { LocalizationHelper } from '@microsoft/mgt';
  LocalizationHelper.strings = {
    _components: {
      'search-box': {
        placeholder: 'Search for 🔥 stuff',
        title: 'Search for content'
      },
    }
  }
  </script>
`;

export const events = () => html`
  <!-- Open dev console and change the search box value -->
  <!-- See js tab for event subscription -->

  <mgt-search-box></mgt-search-box>
  <mgt-search-results entity-types="driveItem"></mgt-search-results>
  <script>
    import '@microsoft/mgt-components/dist/es6/components/preview';
    const searchBox = document.querySelector('mgt-search-box');
    const searchResults = document.querySelector('mgt-search-results');
    searchBox.addEventListener('searchTermChanged', (e) => {
      searchResults.queryString = e.detail;
    });
  </script>
`;
