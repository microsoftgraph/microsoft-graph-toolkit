/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Preview / mgt-search-results / Style',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <style>
    .search-results {
      --answer-border-radius: 10px;
      --answer-box-shadow: 0px 2px 30px pink;;
      --answer-border: dotted 2px white;;
      --answer-padding: 8px 0px;
    }
  </style>
  <mgt-search-results class="search-results" query-string="yammer" entity-types="bookmark" version="beta"></mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;
