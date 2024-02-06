/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker',
  component: 'people-picker',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-people-picker></mgt-people-picker>' },
      editor: {
        hidden: true
      }
    }
  }
};

export const peoplePicker = () => html`
   <mgt-people-picker></mgt-people-picker>
 `;
