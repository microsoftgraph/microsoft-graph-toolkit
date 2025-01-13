/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-taxonomy-picker / HTML',
  component: 'taxonomy-picker',
  decorators: [withCodeEditor]
};

export const TaxonomyPicker = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
`;

export const events = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
  <!-- Check the console tab for results -->
  <script>
    document.querySelector('mgt-taxonomy-picker').addEventListener('updated', e => {
      console.log('updated', e);
    });
    document.querySelector('mgt-taxonomy-picker').addEventListener('selectionChanged', e => {
      console.log('selected term:', e.detail)
    });
  </script>
`;
