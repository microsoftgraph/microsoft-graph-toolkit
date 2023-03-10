/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components/mgt-file-list-composite',
  component: 'file-list-composite',
  decorators: [withCodeEditor]
};

export const fileComposite = () => html`
  <mgt-file-list-composite></mgt-file-list-composite>
`;

export const BreadcrumbRootName = () => html`
  <mgt-file-list-composite breadcrumb-root-name="Drive"></mgt-file-list-composite>
`;

export const UseGridView = () => html`
  <mgt-file-list-composite breadcrumb-root-name="Drive" use-grid-view></mgt-file-list-composite>
`;

export const events = () => html`
<mgt-file-list-composite></mgt-file-list-composite>
<script>
  const component = document.querySelector('mgt-file-list-composite');
  component.addEventListener('breadcrumbclick', e => {
    console.log('breadcrumbclick', e.detail);
  });
  component.addEventListener('itemClick', e => {
    console.log('itemClick', e.detail);
  });
</script>
`;
