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
  decorators: [withCodeEditor],
  tags: ['autodocs'],
  parameters: {
    docs: {
      source: { code: '<mgt-search-results></mgt-search-results>' }
    }
  }
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
