/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-taxonomy-picker',
  component: 'taxonomy-picker',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: {
        code: '<mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>'
      },
      editor: { hidden: true }
    }
  }
};

export const TaxonomyPicker = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
`;
