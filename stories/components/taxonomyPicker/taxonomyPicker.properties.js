/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-taxonomy-picker / Properties',
  component: 'taxonomy-picker',
  decorators: [withCodeEditor]
};

export const DefaultSelectedTermId = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052" default-selected-term-id="71d47d57-479b-4a8c-80df-697da2d5a2e1"></mgt-taxonomy-picker>
`;

export const Disabled = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052" default-selected-term-id="71d47d57-479b-4a8c-80df-697da2d5a2e1" disabled></mgt-taxonomy-picker>
`;

export const Position = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052" position="above" id="mgt-taxonomy-picker-above"></mgt-taxonomy-picker>
  <style>
    #mgt-taxonomy-picker-above {
      float: left;
      margin-top: 200px;
    }
  </style>
`;

export const Locale = () => html`
  <p>In this example, french terms will be shown if they are present</p>
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052" locale="fr-FR"></mgt-taxonomy-picker>
`;
