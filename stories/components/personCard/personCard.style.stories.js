/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person-card / Style',
  component: 'person-card',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
  <style>
    .person-card {
      --person-card-nav-back-arrow-hover-color: green;
      --person-card-icon-color: black;
      --person-card-line1-font-size: 30px;
      --person-card-line1-font-weight: 800;
      --person-card-line1-line-height: 38px;
      --person-card-line2-font-size: 24px;
      --person-card-line2-font-weight: 600;
      --person-card-line2-line-height: 30px;
      --person-card-line3-font-size: 24px;
      --person-card-line3-font-weight: 300;
      --person-card-line3-line-height: 29px;
      --person-card-avatar-size: 85px;
      --person-card-details-left-spacing: 25px;
      --person-card-avatar-top-spacing: 25px;
      --person-card-details-bottom-spacing: 20px;
      --person-card-background-color: pink;
      --person-card-expanded-background-color-hover: blue;
      --person-card-icon-hover-color: magenta;
      --person-card-show-more-color: blue;
      --person-card-show-more-hover-color: green;
      --person-card-fluent-background-color: yellow;
      --person-card-line1-text-color: purple;
      --person-card-line2-text-color: blue;
      --person-card-line3-text-color: blue;
      --person-card-fluent-background-color-hover: orange;
      --organization-active-org-member-target-background-color: blue;
      --file-list-background-color: pink;
      --file-item-background-color: pink;
      --person-card-base-icons-left-spacing: 110px
    }

  </style>
  <mgt-person-card class="person-card" person-query="me"></mgt-person-card>
`;
