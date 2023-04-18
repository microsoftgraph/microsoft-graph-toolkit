/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-taxonomy-picker / Templating',
  component: 'mgt-taxonomy-picker',
  decorators: [withCodeEditor]
};

export const NoDataTemplate = () => html`
    <mgt-taxonomy-picker
        term-set-id="138a652e-7f23-46f6-b480-13da2308c235"
        term-id="12241156-a7f6-4dfb-8d91-5d9ab596c3c7"
      >
      <template data-type="no-data">
        <div>No terms</div>
      </template>
    </mgt-taxonomy-picker>
  `;

export const ErrorTemplate = () => html`
    <mgt-taxonomy-picker
        term-set-id="138a652e-7f23-46f6-b480-13da2308c2351"
      >
      <template data-type="error">
        <div>Error getting terms</div>
      </template>
    </mgt-taxonomy-picker>
  `;
