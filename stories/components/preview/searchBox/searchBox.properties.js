/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Preview / mgt-search-box / Properties',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const setSearchBoxSearchTerm = () => html`
  <mgt-search-box search-term="contoso">
  </mgt-search-box>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchBoxDebounceDelay = () => html`
  <mgt-search-box debounce-delay="1000">
  </mgt-search-box>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setPlaceholder = () => html`
  <mgt-search-box placeholder="Search for content...">
  </mgt-search-box>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;
