/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { IProvider, LoginType, ProviderState } from './IProvider';
import { Graph } from '../Graph';

import { UserAgentApplication, AuthenticationParameters, AuthResponse, AuthError, Configuration } from 'msal';

export interface MsalConfig {
  clientId: string;
  scopes?: string[];
  authority?: string;
  loginType?: LoginType;
  options?: Configuration;
}

export class MsalProvider extends IProvider {
  private _loginType: LoginType;

  protected _userAgentApplication: UserAgentApplication;

  get provider() {
    return this._userAgentApplication;
  }

  scopes: string[];
  protected clientId: string;

  constructor(config: MsalConfig) {
    super();
    this.initProvider(config);
  }

  private initProvider(config: MsalConfig) {
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;

    let tokenReceivedCallbackFunction = ((response: AuthResponse) => {
      this.tokenReceivedCallback(response);
    }).bind(this);

    let errorReceivedCallbackFunction = ((authError: AuthError, accountState: string) => {
      this.errorReceivedCallback(authError, status);
    }).bind(this);

    if (config.clientId) {
      let msalConfig: Configuration = config.options || { auth: { clientId: config.clientId } };

      msalConfig.auth.clientId = config.clientId;
      msalConfig.cache = msalConfig.cache || {};
      msalConfig.cache.cacheLocation = msalConfig.cache.cacheLocation || 'localStorage';
      msalConfig.cache.storeAuthStateInCookie = msalConfig.cache.storeAuthStateInCookie || true;

      if (config.authority) {
        msalConfig.auth.authority = config.authority;
      }

      this.clientId = config.clientId;

      this._userAgentApplication = new UserAgentApplication(msalConfig);
      this._userAgentApplication.handleRedirectCallback(tokenReceivedCallbackFunction, errorReceivedCallbackFunction);
    } else {
      throw 'clientId must be provided';
    }

    this.graph = new Graph(this);

    this.trySilentSignIn();
  }

  async trySilentSignIn() {
    if (this._userAgentApplication.isCallback(window.location.hash)) {
      return;
    }
    if (this._userAgentApplication.getAccount()) {
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }

  async login(): Promise<void> {
    let loginRequest: AuthenticationParameters = {
      scopes: this.scopes,
      prompt: 'select_account'
    };

    if (this._loginType === LoginType.Popup) {
      let response = await this._userAgentApplication.loginPopup(loginRequest);
      this.setState(response.account ? ProviderState.SignedIn : ProviderState.SignedOut);
    } else {
      this._userAgentApplication.loginRedirect(loginRequest);
    }
  }

  async logout(): Promise<void> {
    this._userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    let scopes = options ? options.scopes || this.scopes : this.scopes;
    let accessToken: string;
    let accessTokenRequest: AuthenticationParameters = {
      scopes: scopes
    };
    try {
      let response = await this._userAgentApplication.acquireTokenSilent(accessTokenRequest);
      accessToken = response.accessToken;
    } catch (e) {
      console.log(e);
      if (this.requiresInteraction(e)) {
        if (this._loginType == LoginType.Redirect) {
          // check if the user denied the scope before
          if (!this.areScopesDenied(scopes)) {
            this.setRequestedScopes(scopes);
            this._userAgentApplication.acquireTokenRedirect(accessTokenRequest);
          }
        } else {
          try {
            let response = await this._userAgentApplication.acquireTokenPopup(accessTokenRequest);
            accessToken = response.accessToken;
          } catch (e) {
            console.log('getaccesstoken catch2 : ' + e);
          }
        }
      } else {
        // if we don't know what the error is, just ask the user to sign in again
        this.setState(ProviderState.SignedOut);
      }
    }
    return accessToken;
  }

  updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  private requiresInteraction(error) {
    if (!error || !error.errorCode) {
      return false;
    }
    return (
      error.errorCode.indexOf('consent_required') !== -1 ||
      error.errorCode.indexOf('interaction_required') !== -1 ||
      error.errorCode.indexOf('login_required') !== -1
    );
  }

  private tokenReceivedCallback(response: AuthResponse) {
    if (response.tokenType == 'id_token') {
      this.setState(ProviderState.SignedIn);
    }

    this.clearRequestedScopes();
  }

  private errorReceivedCallback(authError: AuthError, accountState: string) {
    console.log('authError: ' + authError + ' accountState ' + accountState);
    let requestedScopes = this.getRequestedScopes();
    if (requestedScopes) {
      this.addDeniedScopes(requestedScopes);
    }

    this.clearRequestedScopes();
  }

  //session storage
  private ss_requested_scopes_key = 'mgt-requested-scopes';
  private ss_denied_scopes_key = 'mgt-denied-scopes';

  private setRequestedScopes(scopes: string[]) {
    if (scopes) {
      sessionStorage.setItem(this.ss_requested_scopes_key, JSON.stringify(scopes));
    }
  }

  private getRequestedScopes() {
    let scopes_str = sessionStorage.getItem(this.ss_requested_scopes_key);
    return scopes_str ? JSON.parse(scopes_str) : null;
  }

  private clearRequestedScopes() {
    sessionStorage.removeItem(this.ss_requested_scopes_key);
  }

  private addDeniedScopes(scopes: string[]) {
    if (scopes) {
      let deniedScopes: string[] = this.getDeniedScopes() || [];
      deniedScopes = deniedScopes.concat(scopes);

      var index = deniedScopes.indexOf('openid');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }

      index = deniedScopes.indexOf('profile');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }
      sessionStorage.setItem(this.ss_denied_scopes_key, JSON.stringify(deniedScopes));
    }
  }

  private getDeniedScopes() {
    let scopes_str = sessionStorage.getItem(this.ss_denied_scopes_key);
    return scopes_str ? JSON.parse(scopes_str) : null;
  }

  private areScopesDenied(scopes: string[]) {
    if (scopes) {
      const deniedScopes = this.getDeniedScopes();
      if (deniedScopes && deniedScopes.filter(s => -1 !== scopes.indexOf(s)).length > 0) {
        return true;
      }
    }

    return false;
  }
}
