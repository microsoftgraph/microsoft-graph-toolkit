/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people / Templating',
  component: 'people',
  decorators: [withCodeEditor]
};

export const DefaultTemplates = () => html`
<style>
  ul {
    list-style-type: none;
    padding: 0;
  }

  li {
    box-shadow: 0 3px 14px rgba(0, 0, 0, 0.3);
    padding: 8px;
    margin: 8px;
  }

  li mgt-person {
    --avatar-size: 42px;
  }
</style>
<mgt-people>
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
</mgt-people>`;

export const PersonTemplate = () => html`
<mgt-people>
  <template data-type="person">
    {{person.displayName}} |
  </template>
</mgt-people>`;

export const OverflowTemplate = () => html`
<mgt-people>
  <template data-type="overflow">
    and {{extra}} left
  </template>
</mgt-people>`;
