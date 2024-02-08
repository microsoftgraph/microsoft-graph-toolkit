/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-theme-toggle / React',
  component: 'theme-toggle',
  decorators: [
    withCodeEditor({
      disableThemeToggle: true
    })
  ]
};

export const UserPreferenceDriven = () => html`
  <mgt-theme-toggle></mgt-theme-toggle>
  <react>
    import { ThemeToggle } from '@microsoft/mgt-react';

    export default () => (
      <ThemeToggle></ThemeToggle>
    );
  </react>
  <style>
    body {
      background-color: var(--fill-color);
      color: var(--neutral-foreground-rest);
      font-family: var(--body-font);
    }
</style>
`;

export const SetMode = () => html`
  <mgt-theme-toggle mode="light"></mgt-theme-toggle>
  <react>
    import { ThemeToggle } from '@microsoft/mgt-react';

    export default () => (
      <ThemeToggle mode="light"></ThemeToggle>
    );
  </react>
  <style>
    body {
      background-color: var(--fill-color);
    }
  </style>
`;
