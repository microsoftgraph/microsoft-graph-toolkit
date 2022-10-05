/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { IProvider, LoginType, ProviderState, createFromProvider } from '@microsoft/mgt-element';
import { AuthenticationParameters, AuthError, AuthResponse, Configuration, UserAgentApplication } from 'msal';

/**
 * base config for MSAL authentication
 *
 * @export
 * @interface MsalConfigBase
 */
interface MsalConfigBase {
  /**
   * scopes
   *
   * @type {string[]}
   * @memberof MsalConfigBase
   */
  scopes?: string[];
  /**
   * loginType if login uses popup
   *
   * @type {LoginType}
   * @memberof MsalConfigBase
   */
  loginType?: LoginType;
  /**
   * login hint value
   *
   * @type {string}
   * @memberof MsalConfigBase
   */
  loginHint?: string;
  /**
   * Domain hint value
   *
   * @type {string}
   * @memberof MsalConfigBase
   */
  domainHint?: string;
  /**
   * prompt value
   *
   * @type {string}
   * @memberof MsalConfigBase
   */
  prompt?: string;
}

/**
 * config for MSAL authentication where a UserAgentApplication already exists
 *
 * @export
 * @interface MsalConfig
 */
export interface MsalUserAgentApplicationConfig extends MsalConfigBase {
  /**
   * UserAgentApplication instance to use
   *
   * @type {UserAgentApplication}
   * @memberof MsalConfig
   */
  userAgentApplication: UserAgentApplication;
}

/**
 * config for MSAL authentication
 *
 * @export
 * @interface MsalConfig
 */
export interface MsalConfig extends MsalConfigBase {
  /**
   * clientId alphanumeric code
   *
   * @type {string}
   * @memberof MsalConfig
   */
  clientId: string;

  /**
   * config authority
   *
   * @type {string}
   * @memberof MsalConfig
   */
  authority?: string;
  /**
   * options as defined in
   * https://learn.microsoft.com/azure/active-directory/develop/msal-js-initializing-client-applications#configuration-options
   *
   * @type {Configuration}
   * @memberof MsalConfig
   */
  options?: Configuration;
  /**
   * redirect Uri
   *
   * @type {string}
   * @memberof MsalConfig
   */
  redirectUri?: string;
}

/**
 * Msal Provider using MSAL.js to aquire tokens for authentication
 *
 * @export
 * @class MsalProvider
 * @extends {IProvider}
 */
export class MsalProvider extends IProvider {
  /**
   * authentication parameter
   *
   * @type {string[]}
   * @memberof MsalProvider
   */
  public scopes: string[];

  /**
   * Gets the user agent application instance
   *
   * @protected
   * @type {UserAgentApplication}
   * @memberof MsalProvider
   */
  public get userAgentApplication() {
    return this._userAgentApplication;
  }

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtMsalProvider';
  }

  /**
   * client-id authentication
   *
   * @protected
   * @type {string}
   * @memberof MsalProvider
   */
  protected clientId: string;

  private _userAgentApplication: UserAgentApplication;
  private _loginType: LoginType;
  private _loginHint: string;
  private _domainHint: string;
  private _prompt: string;

  // session storage
  private sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
  private sessionStorageDeniedScopesKey = 'mgt-denied-scopes';

  constructor(config: MsalConfig | MsalUserAgentApplicationConfig) {
    super();
    this.initProvider(config);
  }

  /**
   * attempts to sign in user silently
   *
   * @returns
   * @memberof MsalProvider
   */
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

  /**
   * sign in user
   *
   * @param {AuthenticationParameters} [authenticationParameters]
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  public async login(authenticationParameters?: AuthenticationParameters): Promise<void> {
    let loginRequest: AuthenticationParameters = authenticationParameters || {
      loginHint: this._loginHint,
      scopes: this.scopes
    };

    this._prompt ? (loginRequest.prompt = this._prompt) : '';
    this._domainHint ? (loginRequest.extraQueryParameters = { domain_hint: this._domainHint }) : '';

    if (this._loginType === LoginType.Popup) {
      const response = await this._userAgentApplication.loginPopup(loginRequest);
      this.setState(response.account ? ProviderState.SignedIn : ProviderState.SignedOut);
    } else {
      this._userAgentApplication.loginRedirect(loginRequest);
    }
  }

  /**
   * sign out user
   *
   * @returns {Promise<void>}
   * @memberof MsalProvider
   */
  public async logout(): Promise<void> {
    this._userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  /**
   * returns an access token for scopes
   *
   * @param {AuthenticationProviderOptions} options
   * @returns {Promise<string>}
   * @memberof MsalProvider
   */
  public async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest: AuthenticationParameters = {
      loginHint: this._loginHint,
      scopes
    };
    this._domainHint ? (accessTokenRequest.extraQueryParameters = { domain_hint: this._domainHint }) : '';
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

  /**
   * sets scopes
   *
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
  public updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  /**
   * checks if error indicates a user interaction is required
   *
   * @protected
   * @param {*} error
   * @returns
   * @memberof MsalProvider
   */
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

  /**
   * setting scopes in sessionStorage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
  protected setRequestedScopes(scopes: string[]) {
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
  protected getRequestedScopes() {
    const scopesStr = sessionStorage.getItem(this.sessionStorageRequestedScopesKey);
    return scopesStr ? JSON.parse(scopesStr) : null;
  }
  /**
   * clears requested scopes from sessionStorage
   *
   * @protected
   * @memberof MsalProvider
   */
  protected clearRequestedScopes() {
    sessionStorage.removeItem(this.sessionStorageRequestedScopesKey);
  }
  /**
   * sets Denied scopes to sessionStoage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof MsalProvider
   */
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
  protected getDeniedScopes() {
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
  protected areScopesDenied(scopes: string[]) {
    if (scopes) {
      const deniedScopes = this.getDeniedScopes();
      if (deniedScopes && deniedScopes.filter(s => -1 !== scopes.indexOf(s)).length > 0) {
        return true;
      }
    }

    return false;
  }

  private initProvider(config: MsalConfig | MsalUserAgentApplicationConfig) {
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
    this._loginHint = config.loginHint;
    this._domainHint = config.domainHint;
    this._prompt = typeof config.prompt !== 'undefined' ? config.prompt : 'select_account';

    let userAgentApplication: UserAgentApplication;
    let clientId: string;

    if ('clientId' in config) {
      if (config.clientId) {
        const msalConfig: Configuration = config.options || { auth: { clientId: config.clientId } };

        msalConfig.auth.clientId = config.clientId;
        msalConfig.cache = msalConfig.cache || {};
        msalConfig.cache.cacheLocation = msalConfig.cache.cacheLocation || 'localStorage';
        if (
          typeof msalConfig.cache.storeAuthStateInCookie === 'undefined' ||
          msalConfig.cache.storeAuthStateInCookie === null
        ) {
          msalConfig.cache.storeAuthStateInCookie = true;
        }

        if (config.authority) {
          msalConfig.auth.authority = config.authority;
        }

        if (config.redirectUri) {
          msalConfig.auth.redirectUri = config.redirectUri;
        }

        clientId = config.clientId;

        userAgentApplication = new UserAgentApplication(msalConfig);
      } else {
        throw new Error('clientId must be provided');
      }
    } else if ('userAgentApplication' in config) {
      if (config.userAgentApplication) {
        userAgentApplication = config.userAgentApplication;
        const msalConfig = userAgentApplication.getCurrentConfiguration();

        clientId = msalConfig.auth.clientId;
      } else {
        throw new Error('userAgentApplication must be provided');
      }
    } else {
      throw new Error('either clientId or userAgentApplication must be provided');
    }

    this.clientId = clientId;

    this._userAgentApplication = userAgentApplication;
    this._userAgentApplication.handleRedirectCallback(
      response => this.tokenReceivedCallback(response),
      (error, state) => this.errorReceivedCallback(error, state)
    );

    this.graph = createFromProvider(this);

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
