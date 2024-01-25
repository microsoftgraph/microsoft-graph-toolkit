/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-get',
  component: 'get',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: {
        code: `
<mgt-get resource="/me/messages" scopes="mail.read">
  <template>
    <pre>{{ JSON.stringify(value, null, 2) }}</pre>
  </template>
</mgt-get>`
      },
      editor: {
        hidden: true
      }
    }
  }
};

export const Get = () => html`
<mgt-get resource="/me/messages" scopes="mail.read">
  <template>
    <pre>{{ JSON.stringify(value, null, 2) }}</pre>
  </template>
</mgt-get>
`;
