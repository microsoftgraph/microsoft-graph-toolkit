/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';
import { versionInfo } from '../../versionInfo';

export default {
  parameters: {
    version: versionInfo
  },
  title: 'Components / mgt-people-picker / Templating',
  component: 'mgt-people-picker',
  decorators: [withCodeEditor]
};

export const personTemplates = () => html`
<mgt-people-picker>
  <template data-type="selected-person">
		<div>
			🧑 {{person.displayName}}
		</div>
	</template>
  <template data-type="person">
		<div>
			✋ {{person.displayName}} 🤚
		</div>
	</template>
</mgt-people-picker>
`;

export const DefaultTemplates = () => html`
<mgt-people-picker>
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
