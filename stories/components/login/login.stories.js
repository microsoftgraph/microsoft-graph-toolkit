/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { versionInfo } from '../../versionInfo';

export default {
  parameters: {
    version: versionInfo
  },
  title: 'Components / mgt-login',
  component: 'login',
  decorators: [withCodeEditor]
};

export const Login = () => html`
  <mgt-login></mgt-login>
`;

export const CompactLogin = () => html`
  <mgt-login login-view="compact"></mgt-login>
`;

export const AvatarLogin = () => html`
  <mgt-login login-view="avatar"></mgt-login>
`;

export const ShowPresenceLogin = () => html`
  <mgt-login show-presence login-view="full"></mgt-login>
`;

export const Templates = () => html`
  <mgt-login>
    <template data-type="signed-out-button-content">
      👋
    </template>
    <template data-type="signed-in-button-content">
      {{personDetails.givenName}}
    </template>
    <template data-type="flyout-commands">
      <div>
        <button data-props="@click: handleSignOut">Sign Out</button>
        <mgt-person person-query="me" person-card="click">
          <template>
            <button class="profile">
              My Profile
            </button>
          </template>
        </mgt-person>
      </div>
    </template>
      <template data-type="flyout-person-details">
          <div>
              {{personDetails.givenName}}
          </div>
          <div>
              {{personDetails.jobTitle}}
          </div>
          <div>
              {{personDetails.mail}}
          </div>
      </template>
  </mgt-login>

  <style>
  .profile {
      margin-top: 2px;
  }
  </style>
`;

export const RTL = () => html`
  <body dir="rtl">
    <mgt-login></mgt-login>
  </body>
`;

export const Events = () => html`
<mgt-login></mgt-login>
<script>
  const login = document.querySelector('mgt-login');
  login.addEventListener('loginInitiated', (e) => {
    console.log("Login Initiated");
  })
  login.addEventListener('loginCompleted', (e) => {
    console.log("Login Completed");
  })
  login.addEventListener('logoutInitiated', (e) => {
    console.log("Logout Initiated");
  })
  login.addEventListener('logoutCompleted', (e) => {
    console.log("Logout Completed");
  })
</script>
`;

export const localization = () => html`
  <mgt-login></mgt-login>
  <script>
  import { LocalizationHelper } from '@microsoft/mgt';
  LocalizationHelper.strings = {
    _components: {
      login: {
        signInLinkSubtitle: 'Sign In 🤗',
        signOutLinkSubtitle: 'Sign Out 🙋‍♀️'
      },
    }
  }
  </script>
`;

export const MultipleAccounts = () => html`
<mgt-login></mgt-login>
Note: this story configures the MockProvider with data to represent the case of multiple signed in accounts.
It is not possible to sign in with additional accounts or switch the active account.
Please refer to the JavaScript tab if you wish to change which accounts are being show here.
<script>
  import { Providers, MockProvider } from './mgt.storybook.js';
  const signedInAccounts = [{
      name: 'Megan Bowen',
      mail: 'MeganB@M365x214355.onmicrosoft.com',
      id: '48d31887-5fad-4d73-a9f5-3c356e68a038'
    },{
      name: 'Emily Braun',
      mail: 'EmilyB@M365x214355.onmicrosoft.com',
      id: '2804bc07-1e1f-4938-9085-ce6d756a32d2'
    }];
  // initialize the auth provider globally with pre-defined signed in users
  Providers.globalProvider = new MockProvider(true, signedInAccounts);
</script>
`;
