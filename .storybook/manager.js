/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { addons, types } from '@storybook/addons';
import theme from './theme';

import React, { useState } from 'react';
import { useParameter, useChannel } from '@storybook/api';
import { Providers } from '../packages/mgt-element/dist/es6/providers/Providers';
import { ProviderState, LoginType } from '../packages/mgt-element/dist/es6/providers/IProvider';
import { Msal2Provider } from '../packages/providers/mgt-msal2-provider/dist/es6/Msal2Provider';
import { CLIENTID, GETPROVIDER_EVENT, SETPROVIDER_EVENT } from './env';
import { IconButton, Icons } from '@storybook/components';
import { MockProvider } from '@microsoft/mgt-element';
import { PACKAGE_VERSION } from '../packages/mgt-element/dist/es6/utils/version';
import '../packages/mgt-components/dist/es6/components/mgt-login/mgt-login';
import '../packages/mgt-components/dist/es6/components/mgt-person/mgt-person';

const PARAM_KEY = 'signInAddon';
const _allow_signin = true;

document.getElementById('mgt-version').innerText = PACKAGE_VERSION;

const mockProvider = new MockProvider(true);
const msal2Provider = new Msal2Provider({
  clientId: CLIENTID,
  scopes: [
    // capitalize all words in the scope
    'user.read',
    'user.read.all',
    'mail.readBasic',
    'people.read',
    'people.read.all',
    'sites.read.all',
    'user.readbasic.all',
    'contacts.read',
    'presence.read',
    'presence.read.all',
    'tasks.readwrite',
    'tasks.read',
    'calendars.read',
    'group.read.all',
    'files.read',
    'files.read.all',
    'files.readwrite',
    'files.readwrite.all'
  ],
  loginType: LoginType.Popup
});

/*const urlParams = new URLSearchParams(window.location.search);
const canLogin = urlParams.get('login');

if (canLogin === 'true') {
  Providers.globalProvider = msal2Provider;
} else {
  Providers.globalProvider = mockProvider;
}*/

Providers.globalProvider = msal2Provider;

const SignInPanel = () => {
  const [state, setState] = useState(Providers.globalProvider.state);

  const emit = useChannel({
    STORY_RENDERED: id => {
      console.log('storyRendered', id);
    }
  });

  const emitProvider = loginState => {
    if (Providers.globalProvider.state === ProviderState.SignedOut && Providers.globalProvider !== mockProvider) {
      emit(SETPROVIDER_EVENT, { state: loginState, provider: mockProvider, name: 'MgtMockProvider' });
    } else {
      emit(SETPROVIDER_EVENT, { state: loginState, provider: msal2Provider, name: 'MgtMsal2Provider' });
    }
  };

  Providers.onProviderUpdated(() => {
    setState(Providers.globalProvider.state);
    emitProvider(Providers.globalProvider.state);
  });

  emitProvider(state);
  /*let loggedInStyles =
    Providers.globalProvider.state === ProviderState.SignedIn
      ? { pointerEvents: 'none', cursor: 'default' }
      : { marginTop: '3px' };*/
  return (
    <>
      {Providers.globalProvider.state !== ProviderState.SignedIn ? (
        <mgt-login login-view="compact" style={{ marginTop: '3px' }}></mgt-login>
      ) : (
        <mgt-person person-query="me"></mgt-person>
      )}
      {/*Providers.globalProvider.name !== 'MgtMockProvider' &&
        Providers.globalProvider.state === ProviderState.SignedIn && (
          <>
            <mgt-person person-query="me"></mgt-person>
            <IconButton active={true} title="Sign Out">
              <Icons icon="outline" />
            </IconButton>
          </>
        )*/}
    </>
  );
};

addons.setConfig({
  enableShortcuts: false,
  theme
});

addons.register('microsoft/graph-toolkit', storybookAPI => {
  const render = ({ active }) => <SignInPanel />;

  addons.add('mgt/sign-in', {
    type: types.TOOLEXTRA,
    title: 'Sign In',
    match: ({ viewMode }) => true,
    render
  });
});
