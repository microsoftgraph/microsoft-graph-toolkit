/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { MsalProvider } from './MsalProvider';
import * as microsoftTeams from '@microsoft/teams-js';
import { LoginType, ProviderState } from './IProvider';
import { Configuration, UserAgentApplication } from 'msal';

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

  private set accessToken(value: string) {
    this._accessToken = value;
    sessionStorage.setItem(this._sessionStorageTokenKey, value);
    this.setState(value ? ProviderState.SignedIn : ProviderState.SignedOut);
  }

  private get accessToken() {
    return this._accessToken;
  }

  static async isAvailable(): Promise<boolean> {
    const ms = 1000;
    return Promise.race([
      new Promise<boolean>((resolve, reject) => {
        microsoftTeams.initialize();
        microsoftTeams.getContext(function(context) {
          if (context) {
            resolve(true);
          } else {
            resolve(false);
          }
        });
      }),
      new Promise<boolean>((resolve, reject) => {
        let id = setTimeout(() => {
          clearTimeout(id);
          resolve(false);
        }, ms);
      })
    ]);
  }

  static handleAuth() {
    // we are in popup world now - authenticate and handle it

    var url = new URL(window.location.href);

    if (UserAgentApplication.prototype.isCallback(window.location.hash)) {
      new MsalProvider({
        clientId: sessionStorage.getItem(this._sessionStorageClientIdKey)
      });
      return;
    }

    const clientId = url.searchParams.get('clientId');
    if (!clientId) {
      microsoftTeams.authentication.notifyFailure('no clientId provided');
      return;
    }

    // save clientId to use during redirect
    sessionStorage.setItem(this._sessionStorageClientIdKey, clientId);

    let provider = new MsalProvider({
      clientId: clientId // need to add scopes
    });

    const handleProviderState = async () => {
      // how do we handle when user can't sign in
      // change to promise and return status
      if (provider.state === ProviderState.SignedOut) {
        provider.login();
      } else if (provider.state === ProviderState.SignedIn) {
        try {
          let accessToken = await provider.getAccessTokenForScopes(...provider.scopes);
          microsoftTeams.authentication.notifySuccess(accessToken);
        } catch (e) {
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
    return this._accessToken;
  }
}
