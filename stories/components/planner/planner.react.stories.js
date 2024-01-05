/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-planner / React',
  component: 'planner',
  decorators: [withCodeEditor]
};

export const planner = () => html`
  <mgt-planner></mgt-planner>
<react>
import { Planner } from '@microsoft/mgt-react';

export default () => (
  <Planner></Planner>
);
</react>
`;
