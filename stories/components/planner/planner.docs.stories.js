/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-planner',
  component: 'planner',
  decorators: [withCodeEditor],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-planner></mgt-planner>' },
      editor: { hidden: true }
    }
  }
};

export const tasks = () => html`
  <mgt-planner></mgt-planner>
`;
