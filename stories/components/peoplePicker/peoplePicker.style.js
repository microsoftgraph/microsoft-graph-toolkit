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
    --input-border: 10px 4px 10px 10px; /* sets all input area border */

    /* OR individual input border sides */
    --input-border-bottom: 2px rgba(255, 255, 255, 0.5) solid;
    --input-border-right: 2px rgba(255, 255, 255, 0.5) solid;
    --input-border-left: 2px rgba(255, 255, 255, 0.5) solid;
    --input-border-top: 2px rgba(255, 255, 255, 0.5) solid;

    --input-background-color: purple; /* input area background color */
    --input-border-color--hover: #008394; /* input area border hover color */
    --input-border-color--focus: #0f78d4; /* input area border focus color */

    --dropdown-background-color: lightpink; /* selection area background color */
    --dropdown-item-hover-background: purple; /* person background color on hover */

    --selected-person-background-color: pink; /* person item background color */

    --color-sub1: white;
    --placeholder-color: whitesmoke; /* placeholder text color */
    --placeholder-color--focus: pink; /* placeholder text focus color */
  }
</style>
`;
