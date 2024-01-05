/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-file-list',
  component: 'file-list',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-file-list></mgt-file-list>' },
      editor: {
        hidden: true
      }
    }
  }
};

export const fileList = () => html`
  <mgt-file-list></mgt-file-list>
`;
