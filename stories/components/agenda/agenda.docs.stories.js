/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-agenda',
  component: 'agenda',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-agenda></mgt-agenda>' },
      editor: {
        hidden: true
      }
    }
  }
};

export const Agenda = () => html`
  <mgt-agenda></mgt-agenda>
`;
