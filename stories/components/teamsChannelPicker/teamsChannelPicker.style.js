import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker / Style',
  component: 'teams-channel-picker',
  decorators: [withCodeEditor]
};

export const darkTheme = () => html`
  <mgt-teams-channel-picker class="mgt-dark"></mgt-teams-channel-picker>
  <style>
    body {
      background-color: black;
    }
  </style>
`;

export const customCSSProperties = () => html`
  <style>
    .teams-channel-picker {
      --channel-picker-input-border: 2px rgba(255, 255, 255, 0.5) solid; /* sets all input area border */

      /* OR individual input border sides */
      --channel-picker-input-border-bottom: 2px rgba(255, 255, 255, 0.5) solid;
      --channel-picker-input-border-right: 2px rgba(255, 255, 255, 0.5) solid;
      --channel-picker-input-border-left: 2px rgba(255, 255, 255, 0.5) solid;
      --channel-picker-input-border-top: 2px rgba(255, 255, 255, 0.5) solid;

      --channel-picker-input-background-color: #1f1f1f; /* input area background color */
      --channel-picker-input-border-color-hover: #008394; /* input area border hover color */
      --channel-picker-input-border-color-focus: #0f78d4; /* input area border focus color */

      --channel-picker-dropdown-background-color: #1f1f1f; /* channel background color */
      --channel-picker-dropdown-item-hover-background: #333d47; /* channel or team hover background */
      --channel-picker-dropdown-item-selected-background: #0F78D4; /* selected channel background color */

      --channel-picker-color: white; /* input area border focus color */
      --channel-picker-arrow-fill: #ffffff;
      --channel-picker-placeholder-color: #f1f1f1; /* placeholder text color */
      --channel-picker-placeholder-color-focus: rgba(255, 255, 255, 0.8); /* place holder text focus color */
    }
  </style>
  <mgt-teams-channel-picker class="teams-channel-picker" person-query="me"></mgt-teams-channel-picker>
`;
