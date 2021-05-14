/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / Style',
  component: 'mgt-person',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
   <mgt-person-card person-query="me" class="mgt-dark"></mgt-person-card>
   <style>
     body {
       background-color: black;
     }
   </style>
 `;
