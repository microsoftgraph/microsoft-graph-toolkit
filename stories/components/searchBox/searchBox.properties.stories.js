/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-box / Properties',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const setSearchBoxSearchTerm = () => html`
  <mgt-search-box search-term="contoso"></mgt-search-box>
`;

export const setSearchBoxDebounceDelay = () => html`
  <mgt-search-box debounce-delay="1000"></mgt-search-box>
`;

export const setPlaceholder = () => html`
  <mgt-search-box placeholder="Search for content..."></mgt-search-box>
`;
