/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-login / React',
  component: 'login',
  decorators: [withCodeEditor]
};

export const Login = () => html`
  <mgt-login></mgt-login>
<react>
import { Login } from '@microsoft/mgt-react';

export default () => (
  <Login></Login>
);
</react>
`;

export const CompactLogin = () => html`
  <mgt-login login-view="compact"></mgt-login>
<react>
import { Login } from '@microsoft/mgt-react';

export default () => (
  <Login loginView='compact'></Login>
);
</react>
`;

export const AvatarLogin = () => html`
  <mgt-login login-view="avatar"></mgt-login>
<react>
import { Login } from '@microsoft/mgt-react';

export default () => (
  <Login loginView='avatar'></Login>
);
</react>
`;

export const ShowPresenceLogin = () => html`
  <mgt-login show-presence login-view="full"></mgt-login>
  <react>
import { Login } from '@microsoft/mgt-react';

export default () => (
  <Login showPresence={true} loginView='full'></Login>
);
</react>
`;

export const RightAligned = () => html`
<div class="right">
    <mgt-login login-view="compact"></mgt-login>
</div>
<react>
import { Login } from '@microsoft/mgt-react';
import './App.css';

export default () => (
  <div className="right">
    <Login loginView='compact'></Login>
  </div>
);
</react>
<style>
.right {
  display: flex;
  flex-direction: row;
  justify-content: flex-end;
}
</style>
`;
