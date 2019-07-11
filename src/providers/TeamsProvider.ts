/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { MsalProvider, MsalConfig } from './MsalProvider';
import { LoginType, ProviderState } from './IProvider';
import { Configuration, UserAgentApplication } from 'msal';

declare var microsoftTeams: any;

export interface TeamsConfig {
  clientId: string;
  authPopupUrl: string;
  scopes?: string[];
  msalOptions?: Configuration;
}

export class TeamsProvider extends MsalProvider {
  scopes: string[];
  private _authPopupUrl: string;
  private _accessToken: string;

  private _sessionStorageTokenKey = 'mgt-teamsprovider-accesstoken';
  private static _sessionStorageClientIdKey = 'msg-teamsprovider-clientId';
  private static _sessionStorageScopesKey = 'msg-teamsprovider-scopesstr';
  private static _sessionStorageGrantedScopes = 'msg=teamsprovider-approved-scopes';

  private set accessToken(value: string) {
    this._accessToken = value;
    sessionStorage.setItem(this._sessionStorageTokenKey, value);
    this.setState(value ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  private get accessToken() {
    return this._accessToken;
  }

  static async isAvailable() {
    return !!microsoftTeams;
  }

  static handleAuth() {
    if (!this.isAvailable) {
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsProvider');
      return;
    }

    microsoftTeams.initialize();

    // we are in popup world now - authenticate and handle it

    const url = new URL(window.location.href);

    let clientId = sessionStorage.getItem(this._sessionStorageClientIdKey);
    let scopesStr = sessionStorage.getItem(this._sessionStorageScopesKey);

    if (!clientId) {
      clientId = url.searchParams.get('clientId');
      scopesStr = url.searchParams.get('scopes');

      sessionStorage.setItem(this._sessionStorageClientIdKey, clientId);
      sessionStorage.setItem(this._sessionStorageScopesKey, scopesStr);
    }

    if (!clientId) {
      microsoftTeams.authentication.notifyFailure('no clientId provided');
      return;
    }

    let provider = new MsalProvider({
      clientId: clientId,
      scopes: scopesStr ? scopesStr.split(',') : null,
      options: {
        auth: {
          clientId: clientId,
          redirectUri: url.protocol + '//' + url.host + url.pathname
        }
      }
    });

    if (UserAgentApplication.prototype.isCallback(window.location.hash)) {
      // the page should redirect again
      return;
    }

    const handleProviderState = async () => {
      // how do we handle when user can't sign in
      // change to promise and return status
      if (provider.state === ProviderState.SignedOut) {
        provider.login();
      } else if (provider.state === ProviderState.SignedIn) {
        try {
          let accessToken = await provider.getAccessTokenForScopes(...provider.scopes);
          sessionStorage.removeItem(this._sessionStorageClientIdKey);
          sessionStorage.removeItem(this._sessionStorageScopesKey);
          microsoftTeams.authentication.notifySuccess(accessToken);
        } catch (e) {
          sessionStorage.removeItem(this._sessionStorageClientIdKey);
          sessionStorage.removeItem(this._sessionStorageScopesKey);
          microsoftTeams.authentication.notifyFailure(e);
        }
      }
    };

    provider.onStateChanged(handleProviderState);
    handleProviderState();
  }

  constructor(config: TeamsConfig) {
    super({
      clientId: config.clientId,
      loginType: LoginType.Redirect,
      scopes: config.scopes,
      options: config.msalOptions
    });

    if (!TeamsProvider.isAvailable) {
      console.error('Make sure you have referenced the Microsoft Teams sdk before using the TeamsProvider');
      return;
    }

    this._authPopupUrl = config.authPopupUrl;
    microsoftTeams.initialize();
    this.accessToken = sessionStorage.getItem(this._sessionStorageTokenKey);
  }

  async login(): Promise<void> {
    this.setState(ProviderState.Loading);
    return new Promise((resolve, reject) => {
      microsoftTeams.getContext(context => {
        let url = new URL(this._authPopupUrl, new URL(window.location.href));
        url.searchParams.append('clientId', this.clientId);

        if (this.scopes) {
          url.searchParams.append('scopes', this.scopes.join(','));
        }

        microsoftTeams.authentication.authenticate({
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

  async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    let scopes = options ? options.scopes || this.scopes : this.scopes;
    return this.accessToken;
  }
}
