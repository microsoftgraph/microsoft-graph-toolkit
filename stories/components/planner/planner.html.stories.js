/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-planner / HTML',
  component: 'planner',
  decorators: [withCodeEditor]
};

export const planner = () => html`
  <mgt-planner></mgt-planner>
`;

export const plannerWithGroupId = () => html`
  <mgt-planner group-id="45327068-6785-4073-8553-a750d6c16a45"></mgt-planner>
  <!--
    NOTE: the default sandbox tenant doesn't have the required Tasks.ReadWrite and
    Group.ReadWrite.All permissions. Test this component in your tenant.
  -->
`;

export const events = () => html`
  <mgt-planner></mgt-planner>
  <script>
    const planner = document.querySelector('mgt-planner');
    planner.addEventListener('updated', (e) => {
      console.log("Updated", e);
    });

    planner.addEventListener('taskAdded', (e) => {
      console.log("Tasks added", e);
    });

    planner.addEventListener('taskChanged', (e) => {
      console.log("Tasks changed", e);
    });

    planner.addEventListener('taskClick', (e) => {
      console.log("Task clicked", e);
    });

    planner.addEventListener('taskRemoved', (e) => {
      console.log("Tasks removed", e);
    });

    planner.addEventListener('templateRendered', (e) => {
      console.log("Template Rendered", e);
    });
    
  </script>
`;
