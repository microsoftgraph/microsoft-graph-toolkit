/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / Styles',
  component: 'mgt-people',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
 <mgt-people class="mgt-dark"></mgt-people>
 <style>
   body {
     background-color: black;
   }
 </style>
`;

export const customCssProperties = () => html`
<mgt-people></mgt-people>
<style>
  mgt-people {
    --list-margin: 10px 4px 10px 10px; /* Margin for component */
    --avatar-margin: 20px; /* Margin for each person */
    --color: pink /* Text color */
  }
</style>
`;
