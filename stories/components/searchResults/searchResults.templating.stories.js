/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-results / Templating',
  component: 'search-results',
  decorators: [withCodeEditor]
};

export const DefaultTemplates = () => html`
  <mgt-search-results query-string="contoso" entity-types="site">
    <template data-type="default">
        <div data-for="result in value[0].hitsContainers[0].hits">
        <div class="displayName">{{result.resource.displayName}}</div>
        </div>
    </template>
    <template data-type="loading">
        Loading
    </template>
  </mgt-search-results>
`;

export const noDataTemplate = () => html`
  <div>
    <div>No data template</div>
    <mgt-search-results query-string="xyxy" entity-types="driveItem">
        <template data-type="no-data">
        <div>No results found</div>
        </template>
    </mgt-search-results>
</div>
  `;
