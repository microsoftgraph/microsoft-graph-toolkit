/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withSignIn } from '../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Samples | Simple',
  component: 'mgt-combo',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: {
    signInAddon: {
      test: 'test'
    }
  }
};

export const LoginToShowAgenda = () => html`
  <mgt-login></mgt-login>
  <mgt-agenda></mgt-agenda>
`;
