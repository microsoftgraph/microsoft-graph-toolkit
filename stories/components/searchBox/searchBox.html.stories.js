/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-box / HTML',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const searchBox = () => html`
  <mgt-search-box></mgt-search-box>
`;

export const localization = () => html`
  <mgt-search-box></mgt-search-box>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt-element';
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
