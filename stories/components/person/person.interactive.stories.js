/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';

const PersonArgs = {
  view: 'twoLines',
  personQuery: 'me'
};

const PersonArgTypes = {
  view: {
    options: ['image', 'oneLine', 'twoLines', 'threeLines', 'fourLines'],
    control: { type: 'inline-radio' }
  },
  personCardInteraction: {
    options: ['none', 'hover', 'click'],
    control: { type: 'inline-radio' }
  },
  personQuery: { control: { type: 'text' } },
  showPresence: { control: { type: 'boolean' } }
};

export default {
  title: 'Components / mgt-person / Interactive',
  component: 'person',
  argTypes: PersonArgTypes,
  args: PersonArgs,
  parameters: {
    controls: {
      // this uses the expanded docs from the custom elements manifest to add descriptions to the controls panel
      expanded: true,
      // use include to trim the list of args in the controls panel
      include: Object.keys(PersonArgTypes)
    },
    options: { showPanel: true }
  },
  render: ({ view, personQuery, personCardInteraction, showPresence }) => html`
    <mgt-person 
      view="${ifDefined(view)}" 
      person-query="${ifDefined(personQuery)}"
      person-card="${ifDefined(personCardInteraction)}"
      ?show-presence=${showPresence}
    ></mgt-person>
  `
};

export const interactive = { args: PersonArgs };

export const showPresenceFoo = { args: { ...PersonArgs, showPresence: true } };
