/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-file-list',
  component: 'mgt-file-list',
  decorators: [withCodeEditor]
};

export const simple = () => html`
     <mgt-file-list></mgt-file-list>
  `;

export const RTL = () => html`
     <mgt-file-list dir="rtl"></mgt-file-list>
  `;
