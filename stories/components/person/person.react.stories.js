/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-person / React',
  component: 'person',
  decorators: [withCodeEditor]
};

export const person = () => html`
  <mgt-person person-query="me"></mgt-person>
  <br>
  <mgt-person person-query="me" view="oneLine"></mgt-person>
  <br>
  <mgt-person person-query="me" view="twoLines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="threeLines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="fourLines"></mgt-person>
  <react>
    import { Person } from '@microsoft/mgt-react';

    export default () => (
      <>
        <Person personQuery="me"></Person>
        <br />
        <Person personQuery="me" view={ViewType.oneline}></Person>
        <br />
        <Person personQuery="me" view={ViewType.twolines}></Person>
        <br />
        <Person personQuery="me" view={ViewType.threelines}></Person>
        <br />
        <Person personQuery="me" view={ViewType.fourlines}></Person>
      </>
    );
  </react>
`;

export const events = () => html`
  <div style="margin-bottom: 10px">Click on each line</div>
  <div class="example">
    <mgt-person person-query="me" view="fourlines"></mgt-person>
  </div>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { Person, ViewType, IDynamicPerson } from '@microsoft/mgt-react';

    export default () => {
      const onLineClicked = useCallback((e: CustomEvent<IDynamicPerson>) => {
        console.log(e.detail);
      }, []);

      return (
        <Person
          personQuery="me"
          view={ViewType.fourlines}
          line1clicked={onLineClicked}
          line2clicked={onLineClicked}
          line3clicked={onLineClicked}
          line4clicked={onLineClicked}></Person>
      );
    };
  </react>
`;
