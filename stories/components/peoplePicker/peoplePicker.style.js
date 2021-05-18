/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / Styles',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <mgt-people-picker class="mgt-dark"></mgt-people-picker>`;

export const customCssProperties = () => html`
<mgt-people-picker></mgt-people-picker>
<style>
  mgt-people-picker {
    --input-border: 10px 4px 10px 10px; /* Input section entire border */
    --input-border-top: 20px; /* Input section border top only */
    --input-border-right: 10px; /*  */
    --input-background-color: purple;
    --placeholder-color: whitesmoke;
    --color: whitesmoke; /* Input section border right only */
  }
</style>
`;
