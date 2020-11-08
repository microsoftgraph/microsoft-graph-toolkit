/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React, { useState } from 'react';
import { addons, types } from '@storybook/addons';
import { STORIES_CONFIGURED, STORY_MISSING } from '@storybook/core-events';
import { AddonPanel } from '@storybook/components';
import { useParameter, useChannel } from '@storybook/api';
import { MsalProvider } from '../packages/mgt/dist/commonjs';
import { Providers, LoginType } from '../packages/mgt-element/dist';
import { CLIENTID, GETPROVIDER_EVENT, SETPROVIDER_EVENT } from './env';

const PARAM_KEY = 'signInAddon';
const _allow_signin = false;

const msalProvider = new MsalProvider({
  clientId: CLIENTID,
  loginType: LoginType.Popup
});

Providers.globalProvider = msalProvider;

const SignInPanel = () => {
  const value = useParameter(PARAM_KEY, null);

  const [state, setState] = useState(Providers.globalProvider.state);

  const emit = useChannel({
    STORY_RENDERED: id => {
      console.log('storyRendered', id);
    },
    [GETPROVIDER_EVENT]: params => {
      emitProvider(state);
    }
  });

  const emitProvider = loginState => {
    emit(SETPROVIDER_EVENT, { state: loginState });
  };

  Providers.onProviderUpdated(() => {
    setState(Providers.globalProvider.state);
    emitProvider(Providers.globalProvider.state);
  });

  emitProvider(state);

  return (
    <div>
      {_allow_signin ? (
        <mgt-login />
      ) : (
        'All components are using mock data - sign in function will be available in a future release'
      )}
    </div>
  );
};

addons.register('microsoft/graph-toolkit', storybookAPI => {
  storybookAPI.on(STORIES_CONFIGURED, (kind, story) => {
    if (storybookAPI.getUrlState().path === '/story/*') {
      storybookAPI.selectStory('mgt-login', 'login');
    }
  });
  storybookAPI.on(STORY_MISSING, (kind, story) => {
    storybookAPI.selectStory('mgt-login', 'login');
  });

  const render = ({ active, key }) => (
    <AddonPanel active={active} key={key}>
      <SignInPanel />
    </AddonPanel>
  );

  addons.add('mgt/sign-in', {
    type: types.PANEL,
    title: 'Sign In',
    render,
    paramKey: PARAM_KEY
  });
});
