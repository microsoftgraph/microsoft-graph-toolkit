/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-person-card',
  component: 'mgt-person',
  decorators: [withCodeEditor]
};

export const simple = () => html`
 <mgt-person-card person-query="me"></mgt-person-card>
`;

export const events = () => html`
   <!-- Open dev console and click on an event -->
   <!-- See js tab for event subscription -->
 
   <mgt-person-card person-query="me"></mgt-person-card>
   <script>
     const personCard = document.querySelector('mgt-person-card');
     personCard.addEventListener('expanded', () => {
       console.log("expanded");
     })
   </script>
 `;

export const RTL = () => html`
    <mgt-person-card dir="rtl"></mgt-person-card>
 `;
