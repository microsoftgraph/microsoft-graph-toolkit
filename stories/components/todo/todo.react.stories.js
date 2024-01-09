/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-todo / React',
  component: 'todo',
  decorators: [withCodeEditor]
};

export const todos = () => html`
  <mgt-todo></mgt-todo>
  <react>
    import { Todo } from '@microsoft/mgt-react';

    export default () => (
      <Todo></Todo>
    );
  </react>
`;
