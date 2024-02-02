/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';

const PersonArgs = {
  'avatar-size': 'auto',
  view: 'twolines',
  'person-query': 'me',
  'person-card': 'none',
  'show-presence': true
};

const PersonArgTypes = {
  'avatar-size': {
    options: ['auto', 'small', 'large'],
    control: { type: 'select' }
  },
  view: {
    options: ['image', 'oneline', 'twolines', 'threelines', 'fourlines'],
    control: { type: 'inline-radio' }
  },
  'person-card': {
    options: ['none', 'hover', 'click'],
    control: { type: 'inline-radio' }
  },
  'person-query': { control: { type: 'text' } },
  'show-presence': {
    control: { type: 'boolean' }
  }
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
  render: ({
    'avatar-size': avatarSize,
    view,
    'person-query': personQuery,
    'person-card': personCardInteraction,
    'show-presence': showPresence
  }) => html`
    <mgt-person
      avatar-size="${ifDefined(avatarSize)}"
      view="${ifDefined(view)}" 
      person-query="${ifDefined(personQuery)}"
      person-card="${ifDefined(personCardInteraction)}"
      ?show-presence=${showPresence}
    ></mgt-person>
  `
};

export const interactive = { args: PersonArgs };
