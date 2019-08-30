/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Configuration, UserAgentApplication } from 'msal';
import { LoginType, ProviderState } from './IProvider';
import { MsalConfig, MsalProvider } from './MsalProvider';

declare var microsoftTeams: any;

declare global {
  interface Window {
    nativeInterface: any;
  }
}

export interface TeamsConfig {
  clientId: string;
  authPopupUrl: string;
  scopes?: string[];
  msalOptions?: Configuration;
}

export class TeamsProvider extends MsalProvider {
  private set accessToken(value: string) {
    this._accessToken = value;
    if (value) {
      sessionStorage.setItem(TeamsProvider._sessionStorageTokenKey, value);
      this.setState(ProviderState.SignedIn);
    } else {
      sessionStorage.removeItem(TeamsProvider._sessionStorageTokenKey);
      this.setState(ProviderState.SignedOut);
    }
  }

  private get accessToken() {
    return this._accessToken;
  }

  public static microsoftTeamsLib;

  public static get isAvailable() {
    if (window.parent === window.self && window.nativeInterface) {
      // In Teams mobile client
      return true;
    } else if (window.name === 'embedded-page-container' || window.name === 'extension-tab-frame') {
      // In Teams web/desktop client
      return true;
    } else {
      return false;
    }
  }

  public static async handleAuth() {
    // we are in popup world now - authenticate and handle it
    if (!this.isAvailable) {
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsProvider');
      return;
    }

    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;
    teams.initialize();

    // msal checks for the window.opener.msal to check if this is a popup authentication
    // and gets a false positive since teams opens a popup for the authentication.
    // in reality, we are doing a redirect authentication and need to act as if this is the
    // window initiating the authentication
    if (window.opener) {
      window.opener.msal = null;
    }

    const url = new URL(window.location.href);

    let clientId = sessionStorage.getItem(this._sessionStorageClientIdKey);
    let scopesStr = sessionStorage.getItem(this._sessionStorageScopesKey);
    let loginHint = sessionStorage.getItem(this._sessionStorageLoginHint);

    if (!clientId) {
      clientId = url.searchParams.get('clientId');
      scopesStr = url.searchParams.get('scopes');
      loginHint = url.searchParams.get('loginHint');

      sessionStorage.setItem(this._sessionStorageClientIdKey, clientId);
      sessionStorage.setItem(this._sessionStorageScopesKey, scopesStr);
      sessionStorage.setItem(this._sessionStorageLoginHint, loginHint);
    }

    if (!clientId) {
      teams.authentication.notifyFailure('no clientId provided');
      return;
    }

    const scopes = scopesStr ? scopesStr.split(',') : null;

    const provider = new MsalProvider({
      clientId,
      scopes,
      options: {
        auth: {
          clientId,
          redirectUri: url.protocol + '//' + url.host + url.pathname
        },
        system: {
          loadFrameTimeout: 10000
        }
      }
    });

    if ((UserAgentApplication.prototype as any).urlContainsHash(window.location.hash)) {
      // the page should redirect again
      return;
    }

    const handleProviderState = async () => {
      // how do we handle when user can't sign in
      // change to promise and return status
      if (provider.state === ProviderState.SignedOut) {
        provider.login({
          loginHint,
          scopes: scopes || provider.scopes
        });
      } else if (provider.state === ProviderState.SignedIn) {
        try {
          const accessToken = await provider.getAccessTokenForScopes(...provider.scopes);
          sessionStorage.removeItem(this._sessionStorageClientIdKey);
          sessionStorage.removeItem(this._sessionStorageScopesKey);
          sessionStorage.removeItem(this._sessionStorageLoginHint);
          teams.authentication.notifySuccess(accessToken);
        } catch (e) {
          sessionStorage.removeItem(this._sessionStorageClientIdKey);
          sessionStorage.removeItem(this._sessionStorageScopesKey);
          sessionStorage.removeItem(this._sessionStorageLoginHint);
          teams.authentication.notifyFailure(e);
        }
      }
    };

    provider.onStateChanged(handleProviderState);
    handleProviderState();
  }

  private static _sessionStorageTokenKey = 'mgt-teamsprovider-accesstoken';
  private static _sessionStorageClientIdKey = 'msg-teamsprovider-clientId';
  private static _sessionStorageScopesKey = 'msg-teamsprovider-scopesstr';
  private static _sessionStorageLoginHint = 'msg-teamsprovider-loginHint';
  private static _sessionStorageGrantedScopes = 'msg=teamsprovider-approved-scopes';
  private _authPopupUrl: string;
  private _accessToken: string;
  public scopes: string[];

  constructor(config: TeamsConfig) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      scopes: config.scopes,
      options: config.msalOptions
    });

    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;

    this._authPopupUrl = config.authPopupUrl;
    teams.initialize();
    this.accessToken = sessionStorage.getItem(TeamsProvider._sessionStorageTokenKey);
  }

  public async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    const teams = TeamsProvider.microsoftTeamsLib || microsoftTeams;

    return new Promise((resolve, reject) => {
      teams.getContext(context => {
        const url = new URL(this._authPopupUrl, new URL(window.location.href));
        url.searchParams.append('clientId', this.clientId);

        if (context.loginHint) {
          url.searchParams.append('loginHint', context.loginHint);
        }

        if (this.scopes) {
          url.searchParams.append('scopes', this.scopes.join(','));
        }

        teams.authentication.authenticate({
          url: url.href,
          successCallback: result => {
            this.accessToken = result;
            resolve();
          },
          failureCallback: reason => {
            this.accessToken = null;
            reject();
          }
        });
      });
    });
  }

  public async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    if (!this.accessToken) {
      throw null;
    } else {
      return this.accessToken;
    }
  }
}
