/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-people-picker / React',
  component: 'people-picker',
  decorators: [withCodeEditor]
};

export const peoplePicker = () => html`
   <mgt-people-picker></mgt-people-picker>
<react>
import { PeoplePicker } from '@microsoft/mgt-react';

export default () => (
  <PeoplePicker></PeoplePicker>
);
</react>
 `;

export const selectionChangedEvent = () => html`
   <mgt-people-picker></mgt-people-picker>
<react>
import { useCallback } from 'react';
import { PeoplePicker, IDynamicPerson } from '@microsoft/mgt-react';

export default () => {
  const onSelectionChanged = useCallback((e: CustomEvent<IDynamicPerson[]>) => {
    console.log(e.detail);
  }, []);

  return <PeoplePicker selectionChanged={onSelectionChanged}></PeoplePicker>;
</react>
 `;
