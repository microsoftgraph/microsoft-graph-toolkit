/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Editor',
  decorators: [withCodeEditor],
  parameters: {
    viewMode: 'story'
  }
};

export const Editor = () => html`
  <!-- Add your own HTML code here -->

  <script>
    // Add your own JavaScript code here
  </script>
  <style>
    /* Add your own CSS code here */
  </style>
`;
