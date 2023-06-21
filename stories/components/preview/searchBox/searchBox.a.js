/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../../.storybook/addons/codeEditorAddon/codeAddon';
import { defaultDocsPage } from '../../../../.storybook/story-elements/defaultDocsPage';

export default {
  title: 'Preview / mgt-search-box',
  component: 'search-box',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: { code: '<mgt-search-box></mgt-search-box>' }
    }
  }
};

export const searchBox = () => html`
  <mgt-search-box></mgt-search-box>
`;

export const localization = () => html`
  <mgt-search-box></mgt-search-box>
  <script>
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
    const searchBox = document.querySelector('mgt-search-box');
    const searchResults = document.querySelector('mgt-search-results');
    searchBox.addEventListener('searchTermChanged', (e) => {
      searchResults.queryString = e.detail;
    });
  </script>
`;
