import { html } from 'lit-element';
import { withA11y } from '@storybook/addon-a11y';
import { withKnobs } from '@storybook/addon-knobs';
import { withWebComponentsKnobs } from 'storybook-addon-web-components-knobs';
import '../dist/es6/components/mgt-person/mgt-person';
import '../dist/es6/mock/mgt-mock-provider';
import '../dist/es6/mock/MockProvider';

export default {
  title: 'mgt-person',
  component: 'mgt-person',
  decorators: [withA11y, withKnobs, withWebComponentsKnobs],
  parameters: { options: { selectedPanel: 'storybookjs/knobs/panel' } }
};

export const person = () => html`
  <mgt-mock-provider></mgt-mock-provider>
  <mgt-person person-query="me" show-name show-email></mgt-person>
`;
