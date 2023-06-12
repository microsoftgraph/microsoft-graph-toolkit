/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Preview / mgt-search-results / Properties',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const setSearchResultsQueryString = () => html`
  <mgt-search-results query-string="contoso">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsQueryTemplate = () => html`
  <mgt-search-results version="beta" query-string="contoso" query-template="({searchTerms}) Title:Northwind">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsEntityTypes = () => html`
  <style>
    .example {
       margin-bottom: 20px;
     }
  </style>

  <div class="example">
	  <h3>Sites</h3>
    <mgt-search-results query-string="contoso" entity-types="site">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Drives</h3>
    <mgt-search-results query-string="contoso" entity-types="drive">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Drive Items</h3>
    <mgt-search-results query-string="contoso" entity-types="driveItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Lists</h3>
    <mgt-search-results query-string="contoso" entity-types="list">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>List Items</h3>
    <mgt-search-results query-string="contoso" entity-types="listItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Messages</h3>
    <mgt-search-results query-string="marketing" entity-types="message">
    </mgt-search-results>
  </div>
  
  <div class="example">
	  <h3>Events</h3>
    <mgt-search-results query-string="marketing" entity-types="event">
    </mgt-search-results>
  </div>

  
  <div class="example">
	  <h3>Chat Messages</h3>
    <mgt-search-results query-string="marketing" entity-types="chatMessage">
    </mgt-search-results>
  </div>
  
  <div class="example">
	  <h3>Persons</h3>
    <mgt-search-results query-string="bowen" version="beta" entity-types="person">
    </mgt-search-results>
  </div>
  
  <div class="example">
	  <h3>External Items</h3>
    <mgt-search-results query-string="contoso" content-sources="contosoproducts" version="beta" entity-types="externalItem">
    </mgt-search-results>
  </div>

  <div class="example">
	  <h3>Q&A</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="qna">
    </mgt-search-results>
  </div>

  
  <div class="example">
	  <h3>Bookmarks</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="bookmark">
    </mgt-search-results>
  </div>

  
  <div class="example">
	  <h3>Acronyms</h3>
    <mgt-search-results query-string="contoso" version="beta" entity-types="acronym">
    </mgt-search-results>
  </div>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsEntityTypesCombined = () => html`
  <mgt-search-results query-string="contoso" entity-types="driveItem,listItem">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsScopes = () => html`
  <mgt-search-results query-string="contoso" scopes="User.Read.All,Files.Read.All">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsContentSources = () => html`
  <mgt-search-results query-string="contoso" entity-types="externalItem" content-sources="/external/connections/contosoProducts" scopes="ExternalItem.Read.All">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsVersion = () => html`
  <mgt-search-results query-string="contoso" entity-types="bookmark" version="beta" scopes="Bookmark.Read.All">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsSize = () => html`
  <mgt-search-results query-string="contoso" size="20">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsPagingMax = () => html`
  <mgt-search-results query-string="contoso" paging-max="10">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsFetchThumbnail = () => html`
  <mgt-search-results query-string="contoso" fetch-thumbnail>
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsFields = () => html`
  <mgt-search-results query-string="contoso" version="beta" entity-types="driveItem" fields="Title,ID,ContentTypeId">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsEnableTopResults = () => html`
  <mgt-search-results query-string="marketing" entity-types="message" enable-top-results scopes="Mail.Read">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsCacheEnabled = () => html`
  <mgt-search-results query-string="contoso" cache-enabled>
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;

export const setSearchResultsCacheInvalidationPeriod = () => html`
  <mgt-search-results query-string="contoso" cache-enabled cache-invalidation-period="30000">
  </mgt-search-results>
  <script type="module">
  import '@microsoft/mgt-components/dist/es6/components/preview';
  </script>
`;
