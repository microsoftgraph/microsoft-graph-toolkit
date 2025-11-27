/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / Styles',
  component: 'people-picker',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
<mgt-people-picker class="people-picker"></mgt-people-picker>
<style>
  .people-picker {
    --people-picker-selected-option-background-color: orange;
    --people-picker-selected-option-highlight-background-color: red;
    --people-picker-dropdown-background-color: blue;
    --people-picker-dropdown-result-background-color: yellow;
    --people-picker-dropdown-result-hover-background-color: gold;
    --people-picker-dropdown-result-focus-background-color: green;
    --people-picker-dropdown-max-height: 100px;
    --people-picker-dropdown-scrollbar: auto;
    --people-picker-no-results-text-color: white;
    --people-picker-input-background: gray;
    --people-picker-input-border-color: yellow;
    --people-picker-input-hover-background: green;
    --people-picker-input-hover-border-color: red;
    --people-picker-input-focus-background: purple;
    --people-picker-input-focus-border-color: orange;

    --people-picker-input-placeholder-focus-text-color: yellow;
    --people-picker-input-placeholder-hover-text-color: gold;
    --people-picker-input-placeholder-text-color: black;
    --people-picker-search-icon-color: yellow;
    --people-picker-remove-selected-close-icon-color: blue;
    --people-picker-font-size: 16px;

    /** You can also change the person tokens **/
    --person-line1-text-color: blue;
    --person-line2-text-color: red;
  }
</style>
`;
