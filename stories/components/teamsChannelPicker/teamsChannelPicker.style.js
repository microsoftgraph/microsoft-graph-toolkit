import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-teams-channel-picker / Style',
  component: 'mgt-teams-channel-picker',
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
    mgt-teams-channel-picker {
      --input-border: 2px rgb(255, 0, 0) solid;

      --input-background-color: #fcc0e5; 
      --input-border-color--hover: #008394; 
      --input-border-color--focus: #0f78d4; 

      --dropdown-background-color: #FF69B4; 
      --dropdown-item-hover-background: #ff92e6; 
      --dropdown-item-selected-background: #a10980; 

      --color: blue; 
      --arrow-fill: #ffffff;
      --placeholder-color: blue; 
      --placeholder-color--focus: rgba(255, 255, 255, 0.8); 
    }
  </style>
  <mgt-teams-channel-picker person-query="me"></mgt-teams-channel-picker>
`;
