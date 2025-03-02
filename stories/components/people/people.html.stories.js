/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / HTML',
  component: 'people',
  decorators: [withCodeEditor]
};

export const People = () => html`
  <mgt-people show-max="5"></mgt-people>
`;

export const RTL = () => html`
  <body dir="rtl">
    <mgt-people show-max="5"></mgt-people>
  </body>
`;

export const Events = () => html`
<mgt-people people-queries="Megan Bowen"></mgt-people>
<script>
  const people = document.querySelector('mgt-people');
  people.addEventListener('people-rendered', (e) => {
    console.log("People rendered");
  });
  
  people.addEventListener('updated', (e) => {
    console.log("Updated");
  });
</script>
`;
