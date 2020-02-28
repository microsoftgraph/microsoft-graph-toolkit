/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withSignIn } from '../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples | Simple',
  component: 'mgt-combo',
  decorators: [withSignIn, withCodeEditor],
  parameters: {
    a11y: {
      disabled: true
    },
    signInAddon: {
      test: 'test'
    }
  }
};

export const LoginToShowAgenda = () => html`
  <mgt-login></mgt-login>
  <mgt-agenda></mgt-agenda>
`;
