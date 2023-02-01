/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker / Templating',
  component: 'teams-channel-picker',
  decorators: [withCodeEditor]
};

export const error = () => html`
  <mgt-teams-channel-picker>
    <template data-type="error">
      <p>Sorry, no Teams or Channels were found</p>
    </template>
  </mgt-teams-channel-picker>
 `;

export const loading = () => html`
<mgt-teams-channel-picker>
  <template data-type="loading">
    <p>Loading results....</p>
  </template>
</mgt-teams-channel-picker>
`;
