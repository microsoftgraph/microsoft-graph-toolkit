/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const searchResults = () => html`
  <mgt-search-results 
    entity-types="driveItem" 
    fetch-thumbnail="true"
    scopes="Files.Read.All"
    query-string="contoso">
  </mgt-search-results>
`;
