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

export const events = () => html`
  <mgt-todo></mgt-todo>
  <react>
    import { Todo } from '@microsoft/mgt-react';

    export default () => {
    const onUpdated = useCallback((e: CustomEvent<undefined>) => {
      console.log('updated', e);
    });

    const onTemplateRendered = useCallback((e: CustomEvent<MgtElement.TemplateRenderedData>) => {
      console.log('templateRendered', e);
    });

    return (
      <Todo
      updated={onUpdated}
      templateRendered={onTemplateRendered}>
    </Todo>
    );
  };
  </react>
  <script>
    const todo = document.querySelector('mgt-todo');
    todo.addEventListener('updated', (e) => {
      console.log('updated', e);
    });
    todo.addEventListener('templateRendered', (e) => {
      console.log('templateRendered', e);
    });
  </script>
`;
