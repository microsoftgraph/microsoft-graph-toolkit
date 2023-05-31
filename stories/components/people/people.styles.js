/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / Styles',
  component: 'people',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
<mgt-people></mgt-people>
<style>
  .people {
    --people-list-margin: 12px;
    --people-avatar-gap: 8px;
    --people-overflow-font-color: orange;
    --people-overflow-font-size: 16px;
    --people-overflow-font-weight: 600;
  }
</style>
`;
