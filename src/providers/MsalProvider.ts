/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../Graph';
import { IProvider, LoginType, ProviderState } from './IProvider';

import { AuthenticationParameters, AuthError, AuthResponse, Configuration, UserAgentApplication } from 'msal';

export interface MsalConfig {
  clientId: string;
  scopes?: string[];
  authority?: string;
  loginType?: LoginType;
  options?: Configuration;
  loginHint?: string;
}

export class MsalProvider extends IProvider {
  get provider() {
    return this._userAgentApplication;
  }

  public scopes: string[];

  protected _userAgentApplication: UserAgentApplication;
  protected clientId: string;
  private _loginType: LoginType;
  private _loginHint: string;

  // session storage
  private ss_requested_scopes_key = 'mgt-requested-scopes';
  private ss_denied_scopes_key = 'mgt-denied-scopes';

  constructor(config: MsalConfig) {
    super();
    this.initProvider(config);
  }

  public async trySilentSignIn() {
    try {
      if (this._userAgentApplication.isCallback(window.location.hash)) {
        return;
      }
      if (this._userAgentApplication.getAccount() && (await this.getAccessToken(null))) {
        this.setState(ProviderState.SignedIn);
      } else {
        this.setState(ProviderState.SignedOut);
      }
    } catch (e) {
      this.setState(ProviderState.SignedOut);
    }
  }

  public async login(authenticationParameters?: AuthenticationParameters): Promise<void> {
    const loginRequest: AuthenticationParameters = authenticationParameters || {
      loginHint: this._loginHint,
      prompt: 'select_account',
      scopes: this.scopes
    };

    if (this._loginType === LoginType.Popup) {
      const response = await this._userAgentApplication.loginPopup(loginRequest);
      this.setState(response.account ? ProviderState.SignedIn : ProviderState.SignedOut);
    } else {
      this._userAgentApplication.loginRedirect(loginRequest);
    }
  }

  public async logout(): Promise<void> {
    this._userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  public async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest: AuthenticationParameters = {
      loginHint: this._loginHint,
      scopes
    };
    try {
      const response = await this._userAgentApplication.acquireTokenSilent(accessTokenRequest);
      return response.accessToken;
    } catch (e) {
      if (this.requiresInteraction(e)) {
        if (this._loginType === LoginType.Redirect) {
          // check if the user denied the scope before
          if (!this.areScopesDenied(scopes)) {
            this.setRequestedScopes(scopes);
            this._userAgentApplication.acquireTokenRedirect(accessTokenRequest);
          } else {
            throw e;
          }
        } else {
          try {
            const response = await this._userAgentApplication.acquireTokenPopup(accessTokenRequest);
            return response.accessToken;
          } catch (e) {
            throw e;
          }
        }
      } else {
        // if we don't know what the error is, just ask the user to sign in again
        this.setState(ProviderState.SignedOut);
        throw e;
      }
    }
    throw null;
  }

  public updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  protected requiresInteraction(error) {
    if (!error || !error.errorCode) {
      return false;
    }
    return (
      error.errorCode.indexOf('consent_required') !== -1 ||
      error.errorCode.indexOf('interaction_required') !== -1 ||
      error.errorCode.indexOf('login_required') !== -1
    );
  }

  protected setRequestedScopes(scopes: string[]) {
    if (scopes) {
      sessionStorage.setItem(this.ss_requested_scopes_key, JSON.stringify(scopes));
    }
  }

  protected getRequestedScopes() {
    let scopes_str = sessionStorage.getItem(this.ss_requested_scopes_key);
    return scopes_str ? JSON.parse(scopes_str) : null;
  }

  protected clearRequestedScopes() {
    sessionStorage.removeItem(this.ss_requested_scopes_key);
  }

  protected addDeniedScopes(scopes: string[]) {
    if (scopes) {
      let deniedScopes: string[] = this.getDeniedScopes() || [];
      deniedScopes = deniedScopes.concat(scopes);

      let index = deniedScopes.indexOf('openid');
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

  protected getDeniedScopes() {
    let scopes_str = sessionStorage.getItem(this.ss_denied_scopes_key);
    return scopes_str ? JSON.parse(scopes_str) : null;
  }

  protected areScopesDenied(scopes: string[]) {
    if (scopes) {
      const deniedScopes = this.getDeniedScopes();
      if (deniedScopes && deniedScopes.filter(s => -1 !== scopes.indexOf(s)).length > 0) {
        return true;
      }
    }

    return false;
  }

  private initProvider(config: MsalConfig) {
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
    this._loginHint = config.loginHint;

    const tokenReceivedCallbackFunction = ((response: AuthResponse) => {
      this.tokenReceivedCallback(response);
    }).bind(this);

    const errorReceivedCallbackFunction = ((authError: AuthError, accountState: string) => {
      this.errorReceivedCallback(authError, status);
    }).bind(this);

    if (config.clientId) {
      const msalConfig: Configuration = config.options || { auth: { clientId: config.clientId } };

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
      throw new Error('clientId must be provided');
    }

    this.graph = new Graph(this);

    this.trySilentSignIn();
  }

  private tokenReceivedCallback(response: AuthResponse) {
    if (response.tokenType === 'id_token') {
      this.setState(ProviderState.SignedIn);
    }

    this.clearRequestedScopes();
  }

  private errorReceivedCallback(authError: AuthError, accountState: string) {
    const requestedScopes = this.getRequestedScopes();
    if (requestedScopes) {
      this.addDeniedScopes(requestedScopes);
    }

    this.clearRequestedScopes();
  }
}
