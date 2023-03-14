/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-search-box / Style',
  component: 'search-box',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <style>
    mgt-search-box {
      --search-input-width: 200px;
    }
  </style>

  <mgt-search-box></mgt-search-box>
`;
