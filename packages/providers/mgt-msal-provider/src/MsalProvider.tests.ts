/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { it, expect } from '@jest/globals';
import { LoginType } from '@microsoft/mgt-element/src/providers/IProvider';
import { UserAgentApplication } from 'msal';
import { MsalConfig, MsalProvider } from './MsalProvider';

jest.mock('@microsoft/microsoft-graph-client/lib/es/src/Client');
jest.mock('msal/lib-commonjs');

describe('MSALProvider', () => {
  beforeEach(() => {
    jest.resetAllMocks();
    UserAgentApplication.prototype.acquireTokenSilent = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        accessToken: 'abc'
      })
    );
  });

  it('undefined clientId should throw exception', () => {
    const config: MsalConfig = {
      clientId: undefined
    };
    expect(() => {
      new MsalProvider(config);
    }).toThrowError('clientId must be provided');
  });

  it('_loginType Popup should call the loginPopup', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    // tslint:disable-next-line: no-string-literal
    expect(msalProvider['_userAgentApplication'].loginPopup).toBeCalled();
  });

  it('_loginType Redirect should call the loginRedirect', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };
    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    // tslint:disable-next-line: no-string-literal
    expect(msalProvider['_userAgentApplication'].loginRedirect).toBeCalled();
  });

  it('getAccessToken should call acquireTokenSilent', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };
    const msalProvider = new MsalProvider(config);
    const authProviderOptions = {};
    const _ = await msalProvider.getAccessToken(authProviderOptions);
    // tslint:disable-next-line: no-string-literal
    expect(msalProvider['_userAgentApplication'].acquireTokenSilent).toBeCalled();
  });

  it('logout should fire onLoginChanged callback', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    expect.assertions(1);
    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.logout();
  });

  it('login should fire onLoginChanged callback when loginType is Popup', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);

    expect.assertions(1);
    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.login();
  });

  it('login should fire onLoginChanged callback when loginType is Redirect', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };

    UserAgentApplication.prototype.loginRedirect = jest.fn().mockImplementationOnce(() => {
      // msalProvider.tokenReceivedCallback(undefined);
    });

    const msalProvider = new MsalProvider(config);

    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.login();
  });
});
