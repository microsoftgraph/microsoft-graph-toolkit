import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import '../dist/es6/components/mgt-tasks/mgt-tasks';
import '../dist/es6/mock/mgt-mock-provider';
import '../dist/es6/mock/MockProvider';

export default {
  title: 'mgt-tasks',
  component: 'mgt-tasks',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const tasks = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-tasks></mgt-tasks>
`;
