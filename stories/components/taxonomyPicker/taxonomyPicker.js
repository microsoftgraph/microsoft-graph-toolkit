/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { defaultDocsPage } from '../../../.storybook/story-elements/defaultDocsPage';

export default {
  title: 'Components / mgt-taxonomy-picker',
  component: 'taxonomy-picker',
  decorators: [withCodeEditor],
  parameters: {
    docs: {
      page: defaultDocsPage,
      source: {
        code:
          '<mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>'
      }
    }
  }
};

export const TaxonomyPicker = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
`;

export const SelectionChangedEvent = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
  <!-- Check the console tab for results -->
  <script>
    document.querySelector('mgt-taxonomy-picker').addEventListener('selectionChanged', e => {
      console.log('selected term:', e.detail)
    });
  </script>
`;
