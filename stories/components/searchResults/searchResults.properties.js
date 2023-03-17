/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results / Properties',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const setSearchResultsQueryString = () => html`
  <mgt-search-results query-string="contoso">
  </mgt-search-results>
`;

export const setSearchResultsQueryTemplate = () => html`
  <mgt-search-results version="beta" query-string="contoso" query-template="({searchTerms}) Title:Northwind">
  </mgt-search-results>
`;

export const setSearchResultsEntityTypes = () => html`
  <mgt-search-results query-string="contoso" entity-types="driveItem,listItem">
  </mgt-search-results>
`;

export const setSearchResultsScopes = () => html`
  <mgt-search-results query-string="contoso" scopes="User.Read.All,Files.Read.All">
  </mgt-search-results>
`;

export const setSearchResultsContentSources = () => html`
  <mgt-search-results query-string="contoso" entity-types="externalItem" content-sources="contosoProducts">
  </mgt-search-results>
`;

export const setSearchResultsVersion = () => html`
  <mgt-search-results query-string="contoso" entity-types="bookmark" version="beta" scopes="Bookmark.Read.All">
  </mgt-search-results>
`;

export const setSearchResultsFromAndSize = () => html`
  <mgt-search-results query-string="contoso" from="100" size="20">
  </mgt-search-results>
`;

export const setSearchResultsPagingMax = () => html`
  <mgt-search-results query-string="contoso" paging-max="10">
  </mgt-search-results>
`;

export const setSearchResultsFetchThumbnail = () => html`
  <mgt-search-results query-string="contoso" fetch-thumbnail>
  </mgt-search-results>
`;

export const setSearchResultsFields = () => html`
  <mgt-search-results query-string="contoso" entity-types="driveItem" fields="Title,ID,ContentTypeId">
  </mgt-search-results>
`;

export const setSearchResultsEnableTopResults = () => html`
  <mgt-search-results query-string="contoso" entity-types="message" enable-top-results scopes="Mail.Read">
  </mgt-search-results>
`;

export const setSearchResultsCacheEnabled = () => html`
  <mgt-search-results query-string="contoso" cache-enabled>
  </mgt-search-results>
`;

export const setSearchResultsCacheInvalidationPeriod = () => html`
  <mgt-search-results query-string="contoso" cache-enabled cache-invalidation-period="30000">
  </mgt-search-results>
`;
