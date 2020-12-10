/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../packages/mgt-components/dist/es6/components/mgt-login/mgt-login';

export default {
  title: 'Components | mgt-login',
  component: 'mgt-login',
  decorators: [withCodeEditor],
  parameters: {
    options: { selectedPanel: 'mgt/sign-in' },
    signInAddon: {
      test: 'test'
    }
  }
};

export const Login = () => html`
  <mgt-login></mgt-login>
`;

export const Templates = () => html`
  <mgt-login>
    <template data-type="signed-in-button-content">
      {{personDetails.givenName}}
    </template>
    <template data-type="flyout-commands">
      <div>
        <button data-props="@click: handleSignOut">Sign Out</button>
        <button>Go to my profile</button>
      </div>
    </template>
  </mgt-login>
`;

export const darkTheme = () => html`
  <mgt-login class="mgt-dark"></mgt-login>
  <style>
    body {
      background-color: black;
    }
  </style>
`;
