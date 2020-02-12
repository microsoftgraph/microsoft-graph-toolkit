/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../../../dist/es6/Graph';
import { IProvider, LoginType, ProviderState } from '../../../dist/es6/providers/IProvider';
import { AuthenticationParameters, AuthError, AuthResponse, Configuration, UserAgentApplication } from 'msal1.2';

/**
 * Msal Provider using MSAL.js to aquire tokens for authentication
 *
 * @export
 * @class MsalProvider
 * @extends {IProvider}
 */
export class MsalProvider extends IProvider {
  constructor(config) {
    super();
    this.initProvider(config);
    // authentication parameter
    this.scopes = [];

    // Determines application
    this.userAgentApplication = null;

    // client-id authentication
    this.clientId = '';
    this._loginType = null;
    this._loginHint = '';

    // session storage
    this.sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
    this.sessionStorageDeniedScopesKey = 'mgt-denied-scopes';
  }

  /**
   * simplified form of single sign-on (SSO)
   *
   * @returns
   * @memberof MsalProvider
   */
  async trySilentSignIn() {
    try {
      if (this.userAgentApplication.isCallback(window.location.hash)) {
        return;
      }
      if (this.userAgentApplication.getAccount() && (await this.getAccessToken(null))) {
        this.setState(ProviderState.SignedIn);
      } else {
        this.setState(ProviderState.SignedOut);
      }
    } catch (e) {
      this.setState(ProviderState.SignedOut);
    }
  }

  /**
   * login auth Promise, Redirects request, and sets Provider state to SignedIn if response is recieved
   *
   * @param {AuthenticationParameters} [authenticationParameters]
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  async login(authenticationParameters) {
    const loginRequest = authenticationParameters || {
      loginHint: this._loginHint,
      prompt: 'select_account',
      scopes: this.scopes
    };

    if (this._loginType === LoginType.Popup) {
      try {
        const response = await this.userAgentApplication.loginPopup(loginRequest);
        this.setState(response.account ? ProviderState.SignedIn : ProviderState.SignedOut);
      } catch (error) {
        this.setState(ProviderState.SignedOut);
      }
    } else {
      this.userAgentApplication.loginRedirect(loginRequest);
    }
  }

  /**
   * logout auth Promise, sets Provider state to SignedOut
   *
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  async logout() {
    this.userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }
  /**
   * recieves acess token Promise
   *
   * @param {AuthenticationProviderOptions} options
   * @returns {Promise<string>}
   * @memberof MsalProvider
   */
  async getAccessToken(options) {
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest = {
      loginHint: this._loginHint,
      scopes
    };
    try {
      const response = await this.userAgentApplication.acquireTokenSilent(accessTokenRequest);
      return response.accessToken;
    } catch (e) {
      if (this.requiresInteraction(e)) {
        if (this._loginType === LoginType.Redirect) {
          // check if the user denied the scope before
          if (!this.areScopesDenied(scopes)) {
            this.setRequestedScopes(scopes);
            this.userAgentApplication.acquireTokenRedirect(accessTokenRequest);
          } else {
            throw e;
          }
        } else {
          try {
            const response = await this.userAgentApplication.acquireTokenPopup(accessTokenRequest);
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
  /**
   * sets scopes
   *
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
  updateScopes(scopes) {
    this.scopes = scopes;
  }

  /**
   * if login runs into error, require user interaction
   *
   * @protected
   * @param {*} error
   * @returns
   * @memberof MsalProvider
   */
  requiresInteraction(error) {
    if (!error || !error.errorCode) {
      return false;
    }
    return (
      error.errorCode.indexOf('consent_required') !== -1 ||
      error.errorCode.indexOf('interaction_required') !== -1 ||
      error.errorCode.indexOf('login_required') !== -1
    );
  }

  /**
   * setting scopes in sessionStorage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
  setRequestedScopes(scopes) {
    if (scopes) {
      sessionStorage.setItem(this.sessionStorageRequestedScopesKey, JSON.stringify(scopes));
    }
  }

  /**
   * getting scopes from sessionStorage if they exist
   *
   * @protected
   * @returns
   * @memberof MsalProvider
   */
  getRequestedScopes() {
    const scopesStr = sessionStorage.getItem(this.sessionStorageRequestedScopesKey);
    return scopesStr ? JSON.parse(scopesStr) : null;
  }
  /**
   * clears requested scopes from sessionStorage
   *
   * @protected
   * @memberof MsalProvider
   */
  clearRequestedScopes() {
    sessionStorage.removeItem(this.sessionStorageRequestedScopesKey);
  }
  /**
   * sets Denied scopes to sessionStoage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
  addDeniedScopes(scopes) {
    if (scopes) {
      let deniedScopes = this.getDeniedScopes() || [];
      deniedScopes = deniedScopes.concat(scopes);

      let index = deniedScopes.indexOf('openid');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }

      index = deniedScopes.indexOf('profile');
      if (index !== -1) {
        deniedScopes.splice(index, 1);
      }
      sessionStorage.setItem(this.sessionStorageDeniedScopesKey, JSON.stringify(deniedScopes));
    }
  }
  /**
   * gets deniedScopes from sessionStorage
   *
   * @protected
   * @returns
   * @memberof MsalProvider
   */
  getDeniedScopes() {
    const scopesStr = sessionStorage.getItem(this.sessionStorageDeniedScopesKey);
    return scopesStr ? JSON.parse(scopesStr) : null;
  }
  /**
   * if scopes are denied
   *
   * @protected
   * @param {string[]} scopes
   * @returns
   * @memberof MsalProvider
   */
  areScopesDenied(scopes) {
    if (scopes) {
      const deniedScopes = this.getDeniedScopes();
      if (deniedScopes && deniedScopes.filter(s => -1 !== scopes.indexOf(s)).length > 0) {
        return true;
      }
    }

    return false;
  }

  initProvider(config) {
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
    this._loginHint = config.loginHint;

    const tokenReceivedCallbackFunction = (response => {
      this.tokenReceivedCallback(response);
    }).bind(this);

    const errorReceivedCallbackFunction = ((authError, accountState) => {
      this.errorReceivedCallback(authError, status);
    }).bind(this);

    if (config.clientId) {
      const msalConfig = config.options || { auth: { clientId: config.clientId } };

      msalConfig.auth.clientId = config.clientId;
      msalConfig.cache = msalConfig.cache || {};
      msalConfig.cache.cacheLocation = msalConfig.cache.cacheLocation || 'localStorage';
      msalConfig.cache.storeAuthStateInCookie = msalConfig.cache.storeAuthStateInCookie || true;

      if (config.authority) {
        msalConfig.auth.authority = config.authority;
      }

      this.clientId = config.clientId;

      this.userAgentApplication = new UserAgentApplication(msalConfig);
      this.userAgentApplication.handleRedirectCallback(tokenReceivedCallbackFunction, errorReceivedCallbackFunction);
    } else {
      throw new Error('clientId must be provided');
    }

    this.graph = new Graph(this);

    this.trySilentSignIn();
  }

  tokenReceivedCallback(response) {
    if (response.tokenType === 'id_token') {
      this.setState(ProviderState.SignedIn);
    }

    this.clearRequestedScopes();
  }

  errorReceivedCallback(authError, accountState) {
    const requestedScopes = this.getRequestedScopes();
    if (requestedScopes) {
      this.addDeniedScopes(requestedScopes);
    }

    this.clearRequestedScopes();
  }
}
