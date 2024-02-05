/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { ifDefined } from 'lit/directives/if-defined.js';

const objectToStyleRules = obj => {
  return Object.keys(obj)
    .map(key => `${key}: ${obj[key]}`)
    .join(';\n\t');
};

const PersonArgs = {
  view: 'twolines',
  'person-query': 'me',
  'show-presence': true
};

const PersonArgTypes = {
  'avatar-type': {
    options: ['initials', 'photo'],
    control: { type: 'select' }
  },
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
  'person-image': { control: { type: 'text' } },
  'user-id': { control: { type: 'text' } },
  'show-presence': { control: { type: 'boolean' } },
  'vertical-layout': { control: { type: 'boolean' } },
  'fetch-image': { control: { type: 'boolean' } },
  'disable-image-fetch': { control: { type: 'boolean' } },
  'line1-property': { control: { type: 'text' } },
  'line2-property': { control: { type: 'text' } },
  'line3-property': { control: { type: 'text' } },
  'line4-property': { control: { type: 'text' } },
  'person-details': { control: { type: 'object' } },
  'fallback-details': { control: { type: 'object' } },
  'person-presence': { control: { type: 'object' } },
  '--person-background-color': { control: { type: 'color' } },
  '--person-avatar-size': { control: { type: 'text' } },
  '--person-avatar-border': { control: { type: 'text' } },
  '--person-avatar-border-radius': { control: { type: 'text' } },
  '--person-initials-text-color': { control: { type: 'color' } },
  '--person-initials-background-color': { control: { type: 'color' } },
  '--person-line1-font-size': { control: { type: 'text' } },
  '--person-line1-font-weight': { control: { type: 'number', max: 1000, min: 100, step: 100 } },
  '--person-line1-text-color': { control: { type: 'color' } },
  '--person-line1-text-transform': { control: { type: 'text' } },
  '--person-line1-text-line-height': { control: { type: 'text' } },
  '--person-line2-font-size': { control: { type: 'text' } },
  '--person-line2-font-weight': { control: { type: 'number', max: 1000, min: 100, step: 100 } },
  '--person-line2-text-color': { control: { type: 'color' } },
  '--person-line2-text-transform': { control: { type: 'text' } },
  '--person-line2-text-line-height': { control: { type: 'text' } },
  '--person-line3-font-size': { control: { type: 'text' } },
  '--person-line3-font-weight': { control: { type: 'text' } },
  '--person-line3-text-color': { control: { type: 'color' } },
  '--person-line3-text-transform': { control: { type: 'text' } },
  '--person-line3-text-line-height': { control: { type: 'text' } },
  '--person-line4-font-size': { control: { type: 'text' } },
  '--person-line4-font-weight': { control: { type: 'number', max: 1000, min: 100, step: 100 } },
  '--person-line4-text-color': { control: { type: 'color' } },
  '--person-line4-text-transform': { control: { type: 'text' } },
  '--person-line4-text-line-height': { control: { type: 'text' } },
  '--person-details-left-spacing': { control: { type: 'text' } },
  '--person-details-bottom-spacing': { control: { type: 'text' } }
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
      include: Object.keys(PersonArgTypes),
      sort: 'requiredFirst'
    },
    options: { showPanel: true }
  },
  render: ({
    'avatar-type': avatarType,
    'avatar-size': avatarSize,
    view,
    'person-query': personQuery,
    'person-image': personImage,
    'user-id': userId,
    'person-card': personCardInteraction,
    'show-presence': showPresence,
    'vertical-layout': verticalLayout,
    'fetch-image': fetchImage,
    'disable-image-fetch': disableImageFetch,
    'line1-property': line1Property,
    'line2-property': line2Property,
    'line3-property': line3Property,
    'line4-property': line4Property,
    'person-details': personDetails,
    'fallback-details': fallbackDetails,
    'person-presence': personPresence,
    '--person-background-color': personBackgroundColor,
    '--person-avatar-size': personAvatarSize,
    '--person-avatar-border': personAvatarBorder,
    '--person-avatar-border-radius': personAvatarBorderRadius,
    '--person-initials-text-color': personInitialsTextColor,
    '--person-initials-background-color': personInitialsBackgroundColor,
    '--person-line1-font-size': personLine1FontSize,
    '--person-line1-font-weight': personLine1FontWeight,
    '--person-line1-text-color': personLine1TextColor,
    '--person-line1-text-transform': personLine1TextTransform,
    '--person-line1-text-line-height': personLine1TextLineHeight,
    '--person-line2-font-size': personLine2FontSize,
    '--person-line2-font-weight': personLine2FontWeight,
    '--person-line2-text-color': personLine2TextColor,
    '--person-line2-text-transform': personLine2TextTransform,
    '--person-line2-text-line-height': personLine2TextLineHeight,
    '--person-line3-font-size': personLine3FontSize,
    '--person-line3-font-weight': personLine3FontWeight,
    '--person-line3-text-color': personLine3TextColor,
    '--person-line3-text-transform': personLine3TextTransform,
    '--person-line3-text-line-height': personLine3TextLineHeight,
    '--person-line4-font-size': personLine4FontSize,
    '--person-line4-font-weight': personLine4FontWeight,
    '--person-line4-text-color': personLine4TextColor,
    '--person-line4-text-transform': personLine4TextTransform,
    '--person-line4-text-line-height': personLine4TextLineHeight,
    '--person-details-left-spacing': personDetailsLeftSpacing,
    '--person-details-bottom-spacing': personDetailsBottomSpacing
  }) => {
    const styles = {};
    if (personBackgroundColor) {
      styles['--person-background-color'] = personBackgroundColor;
    }
    if (personAvatarSize) {
      styles['--person-avatar-size'] = personAvatarSize;
    }
    if (personAvatarBorder) {
      styles['--person-avatar-border'] = personAvatarBorder;
    }
    if (personAvatarBorderRadius) {
      styles['--person-avatar-border-radius'] = personAvatarBorderRadius;
    }
    if (personInitialsTextColor) {
      styles['--person-initials-text-color'] = personInitialsTextColor;
    }
    if (personInitialsBackgroundColor) {
      styles['--person-initials-background-color'] = personInitialsBackgroundColor;
    }
    if (personLine1FontSize) {
      styles['--person-line1-font-size'] = personLine1FontSize;
    }
    if (personLine1FontWeight) {
      styles['--person-line1-font-weight'] = personLine1FontWeight;
    }
    if (personLine1TextColor) {
      styles['--person-line1-text-color'] = personLine1TextColor;
    }
    if (personLine1TextTransform) {
      styles['--person-line1-text-transform'] = personLine1TextTransform;
    }
    if (personLine1TextLineHeight) {
      styles['--person-line1-text-line-height'] = personLine1TextLineHeight;
    }
    if (personLine2FontSize) {
      styles['--person-line2-font-size'] = personLine2FontSize;
    }
    if (personLine2FontWeight) {
      styles['--person-line2-font-weight'] = personLine2FontWeight;
    }
    if (personLine2TextColor) {
      styles['--person-line2-text-color'] = personLine2TextColor;
    }
    if (personLine2TextTransform) {
      styles['--person-line2-text-transform'] = personLine2TextTransform;
    }
    if (personLine2TextLineHeight) {
      styles['--person-line2-text-line-height'] = personLine2TextLineHeight;
    }
    if (personLine3FontSize) {
      styles['--person-line3-font-size'] = personLine3FontSize;
    }
    if (personLine3FontWeight) {
      styles['--person-line3-font-weight'] = personLine3FontWeight;
    }
    if (personLine3TextColor) {
      styles['--person-line3-text-color'] = personLine3TextColor;
    }
    if (personLine3TextTransform) {
      styles['--person-line3-text-transform'] = personLine3TextTransform;
    }
    if (personLine3TextLineHeight) {
      styles['--person-line3-text-line-height'] = personLine3TextLineHeight;
    }
    if (personLine4FontSize) {
      styles['--person-line4-font-size'] = personLine4FontSize;
    }
    if (personLine4FontWeight) {
      styles['--person-line4-font-weight'] = personLine4FontWeight;
    }
    if (personLine4TextColor) {
      styles['--person-line4-text-color'] = personLine4TextColor;
    }
    if (personLine4TextTransform) {
      styles['--person-line4-text-transform'] = personLine4TextTransform;
    }
    if (personLine4TextLineHeight) {
      styles['--person-line4-text-line-height'] = personLine4TextLineHeight;
    }
    if (personDetailsLeftSpacing) {
      styles['--person-details-left-spacing'] = personDetailsLeftSpacing;
    }
    if (personDetailsBottomSpacing) {
      styles['--person-details-bottom-spacing'] = personDetailsBottomSpacing;
    }
    return html`
<style>
mgt-person {
    ${objectToStyleRules(styles)}
}
</style>
<mgt-person
  avatar-type="${ifDefined(avatarType)}"
  avatar-size="${ifDefined(avatarSize)}"
  view="${ifDefined(view)}" 
  person-query="${ifDefined(personQuery ? personQuery : undefined)}"
  person-image="${ifDefined(personImage ? personImage : undefined)}"
  user-id="${ifDefined(userId ? userId : undefined)}"
  person-card="${ifDefined(personCardInteraction)}"
  ?show-presence=${showPresence}
  ?vertical-layout=${verticalLayout}
  ?fetch-image=${fetchImage}
  ?disable-image-fetch=${disableImageFetch}
  line1-property="${ifDefined(line1Property ? line1Property : undefined)}"
  line2-property="${ifDefined(line2Property ? line2Property : undefined)}"
  line3-property="${ifDefined(line3Property ? line3Property : undefined)}"
  line4-property="${ifDefined(line4Property ? line4Property : undefined)}"
  person-details="${ifDefined(personDetails ? JSON.stringify(personDetails) : undefined)}"
  fallback-details="${ifDefined(fallbackDetails ? JSON.stringify(fallbackDetails) : undefined)}"
  person-presence="${ifDefined(personPresence ? JSON.stringify(personPresence) : undefined)}"
></mgt-person>
`;
  }
};

export const interactive = { args: PersonArgs };
