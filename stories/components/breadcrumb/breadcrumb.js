/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components/mgt-breadcrumb',
  component: 'breadcrumb',
  decorators: [withCodeEditor]
};

export const breadcrumb = () => html`
  <mgt-breadcrumb></mgt-breadcrumb>
  <script>
    const component = document.querySelector('mgt-breadcrumb');
    component.breadcrumb = [
      {
        id: '0',
        name: 'root-item'
      },
      {
        id: '1',
        name: 'node-1'
      },
      {
        id: '2',
        name: 'node-2'
      }
    ];
  </script>
`;

export const events = () => html`
  <mgt-breadcrumb></mgt-breadcrumb>
  <script>
    const component = document.querySelector('mgt-breadcrumb');
    component.breadcrumb = [
      {
        id: '0',
        name: 'root-item'
      },
      {
        id: '1',
        name: 'node-1'
      },
      {
        id: '2',
        name: 'node-2'
      }
    ];
    component.addEventListener('breadcrumbclick', e => {
      console.log('breadcrumbclick', e.detail);
    });
  </script>
`;
