/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-login / Styles',
  component: 'mgt-login',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <mgt-login class="mgt-dark"></mgt-login>
  <style>
    body {
      background-color: black;
    }
  </style>
`;

export const customCssProperties = () => html`
<mgt-login></mgt-login>
<style>
  mgt-login {
    --font-size: 14px;
    --font-weight: 600;
    --weight: '100%';
    --height: '100%';
    --margin: 0;
    --padding: 12px 20px;
    --button-color: #1e2020;
    --button-color--hover: var(--theme-primary-color);
    --button-background-color: pink;
    --button-background-color--hover: ##e9ba0f52;
    --popup-background-color: rgba(131, 180, 228, 0.664);
    --popup-command-font-size: 12px;
    --popup-color: #201f1e;
  }
</style>
`;
