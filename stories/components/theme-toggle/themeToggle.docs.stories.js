/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-theme-toggle',
  component: 'theme-toggle',
  decorators: [
    withCodeEditor({
      disableThemeToggle: true
    })
  ],
  tags: ['autodocs', 'hidden'],
  parameters: {
    docs: {
      source: { code: '<mgt-theme-toggle></mgt-theme-toggle>' },
      editor: { hidden: true }
    }
  }
};

export const userPreferenceDriven = () => html`
  <mgt-theme-toggle></mgt-theme-toggle>
  <style>
    body {
        background-color: var(--fill-color);
        color: var(--neutral-foreground-rest);
        font-family: var(--body-font);
    }
  </style>
`;
