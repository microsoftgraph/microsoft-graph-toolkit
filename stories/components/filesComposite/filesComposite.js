/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components/mgt-file-composite',
  component: 'file-composite',
  decorators: [withCodeEditor]
};

export const fileComposite = () => html`
  <mgt-file-composite></mgt-file-composite>
`;

export const BreadcrumbRootName = () => html`
  <mgt-file-composite breadcrumb-root-name="Drive"></mgt-file-composite>
`;

export const events = () => html`
<!-- Open dev console and click for an event -->
<mgt-file-composite></mgt-file-composite>
<script>
  const component = document.querySelector('mgt-file-composite');
  component.addEventListener('breadcrumbclick', e => {
    console.log('breadcrumbclick', e.detail);
  });
  component.addEventListener('itemClick', e => {
    console.log('itemClick', e.detail);
  });
</script>
`;
