/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-get / React',
  component: 'get',
  decorators: [withCodeEditor]
};

export const Get = () => html`
<mgt-get resource="/me/messages" scopes="mail.read">
  <template>
    <pre>{{ JSON.stringify(value, null, 2) }}</pre>
  </template>
</mgt-get>
  <react>
    import { Get, MgtTemplateProps } from '@microsoft/mgt-react';

    export const Messages = (props: MgtTemplateProps) => {
      const value = props.dataContext;
      return (
        <pre>{ JSON.stringify(value, null, 2) }</pre>
      );
    };

    export default () => (
      <Get resource='/me/messages' scopes={['mail.read']}>
        <Messages template="default"></Messages>
      </Get>
    );
  </react>
`;

export const events = () => html`
  <mgt-get resource="/me/messages" scopes="mail.read">
    <template>
      <pre>{{ JSON.stringify(value, null, 2) }}</pre>
    </template>
  </mgt-get>
  <react>
    import { Get, MgtTemplateProps } from '@microsoft/mgt-react';

    export const Messages = (props: MgtTemplateProps) => {
      const value = props.dataContext;
      return (
        <pre>{ JSON.stringify(value, null, 2) }</pre>
      );
    };

    const onUpdated = (e: CustomEvent<undefined>) => {
      console.log('updated', e);
    };
    const onTemplateRendered = (e: CustomEvent<undefined>) => {
      console.log('templateRendered', e);
    };
    const onDataChange = (e: CustomEvent<undefined>) => {
      console.log('dataChange', e);
    };

    export default () => (
      <Get 
    resource='/me/messages' 
    scopes={['mail.read']}
    updated={onUpdated}
    templateRendered={onTemplateRendered}
    dataChange={onDataChange}>
      <Messages template="default"></Messages>
  </Get>
    );
  </react>

  <script>
    const get = document.querySelector('mgt-get');
    get.addEventListener('updated', e => {
      console.log('updated', e);
    });
    get.addEventListener('dataChange', e => {
      console.log('dataChange', e);
    });
    get.addEventListener('templateRendered', e => {
      console.log('templateRendered', e);
    });
  </script>
`;
