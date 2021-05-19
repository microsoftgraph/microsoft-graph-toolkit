/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components | mgt-login',
  component: 'mgt-login',
  decorators: [withCodeEditor]
};

export const Login = () => html`
  <mgt-login></mgt-login>
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
        <button>Go to my profile</button>
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
