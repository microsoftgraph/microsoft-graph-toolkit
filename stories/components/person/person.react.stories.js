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
  <mgt-person person-query="me" view="oneline"></mgt-person>
  <br>
  <mgt-person person-query="me" view="twolines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="threelines"></mgt-person>
  <br>
  <mgt-person person-query="me" view="fourlines"></mgt-person>
  <react>
    import { Person } from '@microsoft/mgt-react';

    export default () => (
      <>
        <Person personQuery="me"></Person>
        <br />
        <Person personQuery="me" view="oneline"></Person>
        <br />
        <Person personQuery="me" view="twolines"></Person>
        <br />
        <Person personQuery="me" view="threelines"></Person>
        <br />
        <Person personQuery="me" view="fourlines"></Person>
      </>
    );
  </react>
`;

export const personCard = () => html`
   <div class="example">
     <div style="margin-bottom:10px">Person card Hover</div>
     <mgt-person person-query="me" view="twolines" person-card="hover"></mgt-person>
   </div>
   <div class="example">
     <div style="margin-bottom:10px">Person card Click</div>
     <mgt-person person-query="me" view="twolines" person-card="click"></mgt-person>
   </div>

   <react>
    import { Person } from '@microsoft/mgt-react';

    export default () => (
      <>
        <div className="example">
          <div style={{'marginBottom': '10px'}}>Person card Hover</div>
          <Person personQuery="me" view="twolines" personCardInteraction="hover"></Person>
        </div>
        <div className="example">
        <div style={{'marginBottom': '10px'}}>Person card Click</div>
          <Person personQuery="me" view="twolines" personCardInteraction="click"></Person>
        </div>
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
    import { Person, IDynamicPerson } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e: CustomEvent<undefined>) => {
        console.log('component updated', e);
      }, []);

      const onLineClicked = useCallback((e: CustomEvent<IDynamicPerson>) => {
        console.log(e.detail);
      }, []);

      return (
        <Person
        personQuery="me"
        view="fourlines"
        updated={onUpdated}
        line1clicked={onLineClicked}
        line2clicked={onLineClicked}
        line3clicked={onLineClicked}
        line4clicked={onLineClicked}>
    </Person>
      );
    };
  </react>
  <script>
    person.addEventListener('updated', e => {
        console.log('updated', e);
    });
    person.addEventListener('line1clicked', e => {
        console.log(e.detail);
    });
    person.addEventListener('line2clicked', e => {
        console.log(e.detail);
    });
    person.addEventListener('line3clicked', e => {
        console.log(e.detail);
    });
    person.addEventListener('line4clicked', e => {
        console.log(e.detail);
    });
  </script>
`;
