/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withSignIn } from '../../.storybook/addons/signInAddon/signInAddon';
import { withCodeEditor } from '../../.storybook/addons/codeEditorAddon/codeAddon';
import '../../dist/es6/components/mgt-teams-channel-picker/mgt-teams-channel-picker';

export default {
  title: 'Components | mgt-teams-channel-picker',
  component: 'mgt-teams-channel-picker',
  decorators: [withA11y, withSignIn, withCodeEditor],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const teamsChannelPicker = () => html`
  <mgt-teams-channel-picker></mgt-teams-channel-picker>
`;

const darkStyle = `
--input-border: 2px rgba(255, 255, 255, 0.5) solid;
--input-border-bottom: 2px rgba(255, 255, 255, 0.5) solid;
--input-border-right: 2px rgba(255, 255, 255, 0.5) solid;
--input-border-left: 2px rgba(255, 255, 255, 0.5) solid;
--input-border-top: 2px rgba(255, 255, 255, 0.5) solid;
--input-background-color: #1f1f1f;
--selection-background-color: #1f1f1f;
--input-hover-color: #008394;
--input-focus-color: #0f78d4;
--selection-hover-color: #333d47;
--font-color: white;
--arrow-fill: #ffffff;
--placeholder-focus-color: rgba(255, 255, 255, 0.8); 
`;

export const DarkMode = () => html`
  <mgt-teams-channel-picker style=${darkStyle}></mgt-teams-channel-picker>
`;