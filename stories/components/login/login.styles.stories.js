/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-login / Styles',
  component: 'login',
  decorators: [withCodeEditor]
};

export const customCSSProperties = () => html`
<mgt-login class="login"></mgt-login>
<style>
  .login {
    --login-signed-out-button-background: yellow;
    --login-signed-out-button-hover-background: orange;
    --login-signed-out-button-text-color: maroon;
    --login-signed-in-background: yellow;
    --login-signed-in-hover-background: green;
    --login-button-padding:5px;
    --login-popup-background-color: blue;
    --login-popup-command-button-background-color: orange;
    --login-popup-padding: 8px;
    --login-add-account-button-text-color: maroon;
    --login-add-account-button-background-color: yellow;
    --login-add-account-button-hover-background-color: white;
    --login-command-button-background-color: orange;
    --login-command-button-hover-background-color: purple;
    --login-command-button-text-color: black;
    --login-account-item-hover-bg-color: black;
    --login-flyout-command-text-color: maroon;

    /** person component tokens **/
    --person-line1-text-color: maroon;
    --person-line2-text-color: maroon;
    --person-background-color: blue;
  }
</style>
<script>
  import { Providers, MockProvider } from './mgt.storybook.js';
  const signedInAccounts = [{
      name: 'Megan Bowen',
      mail: 'MeganB@M365x214355.onmicrosoft.com',
      id: '48d31887-5fad-4d73-a9f5-3c356e68a038'
  },
  {
      name: 'Emily Braun',
      mail: 'EmilyB@M365x214355.onmicrosoft.com',
      id: '2804bc07-1e1f-4938-9085-ce6d756a32d2'
  },
  {
      name: 'Lynne Robbins',
      mail: 'LynneR@M365x214355.onmicrosoft.com',
      id: 'e8a02cc7-df4d-4778-956d-784cc9506e5a'
  },
  {
      name: 'Henrietta Mueller',
      mail: 'HenriettaM@M365x214355.onmicrosoft.com',
      id: 'c8913c86-ceea-4d39-b1ea-f63a5b675166'
  },
  ];
  // initialize the auth provider globally with pre-defined signed in users
  Providers.globalProvider = new MockProvider(true, signedInAccounts);
</script>
`;
