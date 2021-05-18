/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / Templating',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const DefaultTemplates = () => html`
<mgt-people-picker>
  <template>
    <ul><li data-for="person in people">
    <mgt-person data-props="personDetails: person" fetch-image></mgt-person>
      <h3>{{ person.displayName }}</h3>
      <p>{{ person.jobTitle }}</p>
      <p>{{ person.department }}</p>
    </li></ul>
  </template>
  <template data-type="loading">
		<div class="root">
			loading
		</div>
	</template>
	<template data-type="no-data">
		<div class="root">
			there is no data
		</div>
	</template>
</mgt-people-picker>
`;
