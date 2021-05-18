/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-people-picker',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const peoplePicker = () => html`
  <mgt-people-picker></mgt-people-picker>
`;

export const RTL = () => html`
  <body dir="rtl">
    <mgt-people-picker></mgt-people-picker>
  </body>
`;

export const selectionChangedEvent = () => html`
  <mgt-people-picker></mgt-people-picker>
  <!-- Check the console tab for results -->
  <script>
  document.querySelector('mgt-people-picker').addEventListener('selectionChanged', e => {
    console.log(e.detail)
  });
  </script>
`;

export const localization = () => html`
  <mgt-people-picker></mgt-people-picker>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt';
  LocalizationHelper.strings = {
    _components: {
      'people-picker': {
        inputPlaceholderText: 'Search for 🤼',
        noResultsFound: '🤷‍♀️',
        loadingMessage: '🦔'
      }
    }
  }
  </script>
`;
