/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-search-suggestion/mgt-search-suggestion';

export default {
  title: 'Components | mgt-search-suggestion',
  component: 'mgt-search-suggestion',
  decorators: [withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const searchSuggestion = () => html`
  <mgt-search-suggestion></mgt-search-suggestion>
`;

export const suggestionCount = () => html`
  <mgt-search-suggestion
    max-text-suggestion-count="3"
    max-file-suggestion-count="2"
    max-people-suggestion-count="3"
  ></mgt-search-suggestion>
`;

export const suggestionEntity = () => html`
  <mgt-search-suggestion selected-entity-types="file, text, people"></mgt-search-suggestion>
`;

export const cvid = () => html`
  <mgt-search-suggestion cvid="d8e48cff-9cac-40c9-0b5c-e6d24488f781"></mgt-search-suggestion>
`;

export const textDecorations = () => html`
  <mgt-search-suggestion text-decorations="1"></mgt-search-suggestion>
`;

export const darkTheme = () => html`
  <mgt-search-suggestion class="mgt-dark"></mgt-search-suggestion>
`;

export const callback = () =>
  html`
    <mgt-search-suggestion id="search-suggestion"> </mgt-search-suggestion>

    <script>
      function onClickCallback(suggestionValue) {
        console.log('suggestion value:', suggestionValue);
        var searchValue = getSuggestionValue(suggestionValue);
        window.location.assign('https://www.bing.com/search?q=' + searchValue);
      }

      function onEnterKeyPressCallback(originalValue, selectedSuggestionValue) {
        console.log('original value:', originalValue);
        console.log('suggestion value:', selectedSuggestionValue);
        var searchValue = getSuggestionValue(selectedSuggestionValue);
        window.location.assign('https://www.bing.com/search?q=' + searchValue);
      }

      function getSuggestionValue(suggestionValue) {
        var searchValue = '';
        if (suggestionValue.entity == 'File') {
          searchValue = suggestionValue.name;
        } else if (suggestionValue.entity == 'Text') {
          searchValue = suggestionValue.text;
        } else if (suggestionValue.entity == 'People') {
          searchValue = suggestionValue.displayName;
        }
        return searchValue;
      }

      var obj = document.getElementById('search-suggestion');
      obj.onClickCallback = onClickCallback;
      obj.onEnterKeyPressCallback = onEnterKeyPressCallback;
    </script>
  `;
