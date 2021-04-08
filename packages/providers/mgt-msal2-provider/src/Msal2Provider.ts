import {
  IProvider,
  LoginType,
  ProviderState,
  createFromProvider,
  Providers,
  IProviderAccount
} from '@microsoft/mgt-element';
import {
  Configuration,
  PublicClientApplication,
  SilentRequest,
  PopupRequest,
  RedirectRequest,
  AuthenticationResult,
  AccountInfo,
  EndSessionRequest,
  InteractionRequiredAuthError
} from '@azure/msal-browser';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';

/**
 * Config for MSAL2.0 Authentication
 *
 * @export
 * @interface Msal2Config
 */
export interface Msal2Config {
  /**
   * Client ID of app registration
   *
   * @type {string}
   * @memberof Msal2Config
   */
  clientId: string;

  /**
   * LoginType
   *
   * @type {LoginType}
   * @memberof Msal2Config
   */
  loginType?: LoginType;

  /**
   * List of scopes required
   *
   * @type {string[]}
   * @memberof Msal2Config
   */
  scopes?: string[];

  /**
   * LoginHint
   *
   * @type {string}
   * @memberof Msal2Config
   */
  loginHint?: string;

  /**
   * Session ID
   *
   * @type {string}
   * @memberof Msal2Config
   */
  sid?: string;

  /**
   * Domain hint
   *
   * @type {string}
   * @memberof Msal2Config
   */
  domain_hint?: string;

  /**
   * Redirect URI
   *
   * @type {string}
   * @memberof Msal2Config
   */
  redirectUri?: string;

  /**
   * Authority URL
   *
   * @type {string}
   * @memberof Msal2Config
   */
  authority?: string;

  /**
   * Disable multi account functionality
   *
   * @type {boolean}
   * @memberof Msal2Config
   */
  isMultiAccountDisabled?: boolean;

  /**
   * Other options
   *
   * @type {Configuration}
   * @memberof Msal2Config
   */
  options?: Configuration;
}

/**
 * MSAL2Provider using msal-browser to acquire tokens for authentication
 *
 * @export
 * @class Msal2Provider
 * @extends {IProvider}
 */
export class Msal2Provider extends IProvider {
  private _publicClientApplication: PublicClientApplication;

  /**
   * Login type, Either Redirect or Popup
   *
   * @private
   * @type {LoginType}
   * @memberof Msal2Provider
   */
  private _loginType: LoginType;

  /**
   * Login hint, if provided
   *
   * @private
   * @memberof Msal2Provider
   */
  private _loginHint;

  /**
   * Domain hint if provided
   *
   * @private
   * @memberof Msal2Provider
   */
  private _domainHint;

  /**
   * Session ID, if provided
   *
   * @private
   * @memberof Msal2Provider
   */
  private _sid;

  /**
   * Configuration settings for authentication
   *
   * @private
   * @type {Configuration}
   * @memberof Msal2Provider
   */
  private ms_config: Configuration;

  /**
   * Gets the PublicClientApplication Instance
   *
   * @private
   * @type {PublicClientApplication}
   * @memberof Msal2Provider
   */
  public get publicClientApplication() {
    return this._publicClientApplication;
  }

  /**
   * List of scopes
   *
   * @type {string[]}
   * @memberof Msal2Provider
   */
  public scopes: string[];

  private sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
  private sessionStorageDeniedScopesKey = 'mgt-denied-scopes';
  private homeAccountKey = '275f3731-e4a4-468a-bf9c-baca24b31e26';

  public constructor(config: Msal2Config) {
    super();
    this.initProvider(config);
  }

  /**
   * Initialize provider with configuration details
   *
   * @private
   * @param {Msal2Config} config
   * @memberof Msal2Provider
   */
  private async initProvider(config: Msal2Config) {
    if (config.clientId) {
      const msalConfig: Configuration = config.options || { auth: { clientId: config.clientId } };
      this.ms_config = msalConfig;
      this.ms_config.auth.clientId = config.clientId;
      if (config.authority) {
        this.ms_config.auth.authority = config.authority;
      }
      if (config.redirectUri) {
        this.ms_config.auth.redirectUri = config.redirectUri;
      }

      this.ms_config.cache = msalConfig.cache || {};
      this.ms_config.cache.cacheLocation = msalConfig.cache.cacheLocation || 'localStorage';
      if (
        typeof this.ms_config.cache.storeAuthStateInCookie === 'undefined' ||
        this.ms_config.cache.storeAuthStateInCookie === null
      ) {
        this.ms_config.cache.storeAuthStateInCookie = true;
      }

      if (config.redirectUri) {
        this.ms_config.auth.redirectUri = config.redirectUri;
      }
      this.ms_config.system = msalConfig.system || {};
      this.ms_config.system.iframeHashTimeout = msalConfig.system.iframeHashTimeout || 10000;
      this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
      this._loginHint = typeof config.loginHint !== 'undefined' ? config.loginHint : null;
      this._sid = typeof config.sid !== 'undefined' ? config.sid : null;
      this._domainHint = typeof config.domain_hint !== 'undefined' ? config.domain_hint : null;
      this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
      this._publicClientApplication = new PublicClientApplication(this.ms_config);
      this.isMultipleAccountDisabled =
        typeof config.isMultiAccountDisabled !== 'undefined' ? config.isMultiAccountDisabled : true; //Set this to true for now. Once multi account is enabled, default will be false
      this.graph = createFromProvider(this);
      try {
        const tokenResponse = await this._publicClientApplication.handleRedirectPromise();
        if (tokenResponse !== null) {
          this.handleResponse(tokenResponse);
        } else {
          this.trySilentSignIn();
        }
      } catch (e) {
        throw e;
      }
    } else {
      throw new Error('clientId must be provided');
    }
  }

  /**
   * Attempts to sign in user silently
   *
   * @memberof Msal2Provider
   */
  public async trySilentSignIn() {
    let silentRequest: any = {
      scopes: this.scopes,
      domainHint: this._domainHint
    };
    if (this._sid || this._loginHint) {
      silentRequest.sid = this._sid;
      silentRequest.loginHint = this._loginHint;
    } else {
      const account: any = this.getAccount();
      if (account) {
        silentRequest.sid = account.idTokenClaims.sid || null;
        silentRequest.loginHint = account.idTokenClaims.preferred_username;
      }
    }
    if (silentRequest.sid || silentRequest.loginHint) {
      try {
        this.setState(ProviderState.Loading);
        const response = await this._publicClientApplication.ssoSilent(silentRequest);
        if (response) {
          this.handleResponse(response);
        }
      } catch (e) {
        this.setState(ProviderState.SignedOut);
      }
    } else {
      this.setState(ProviderState.SignedOut);
    }
  }

  /**
   * Log in the user
   *
   * @return {*}  {Promise<void>}
   * @memberof Msal2Provider
   */
  public async login(): Promise<void> {
    const loginRequest: PopupRequest = {
      scopes: this.scopes,
      loginHint: this._loginHint,
      prompt: 'select_account',
      domainHint: this._domainHint
    };
    if (this._loginType == LoginType.Popup) {
      const response = await this._publicClientApplication.loginPopup(loginRequest);
      this.handleResponse(response);
    } else {
      const loginRedirectRequest: RedirectRequest = { ...loginRequest };
      this._publicClientApplication.loginRedirect(loginRedirectRequest);
    }
  }

  /**
   * Get all signed in accounts
   *
   * @return {*}
   * @memberof Msal2Provider
   */
  public getAllAccounts() {
    let usernames = [];
    this._publicClientApplication.getAllAccounts().forEach((account: AccountInfo) => {
      usernames.push({ username: account.username, id: account.homeAccountId });
    });
    return usernames;
  }

  /**
   * Switching between accounts
   *
   * @param {*} user
   * @memberof Msal2Provider
   */
  public setActiveAccount(user: IProviderAccount) {
    this._publicClientApplication.setActiveAccount(this._publicClientApplication.getAccountByHomeId(user.id));
    this.setStoredAccount();
    super.setActiveAccount(user);
  }

  /**
   * Once a succesful login occurs, set the active account and store it
   *
   * @param {(AuthenticationResult | null)} response
   * @memberof Msal2Provider
   */
  handleResponse(response: AuthenticationResult | null) {
    if (response !== null) {
      this.setActiveAccount({
        username: response.account.name,
        id: response.account.homeAccountId
      } as IProviderAccount);
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
    this.clearRequestedScopes();
  }

  /**
   * Store the currently signed in account in storage
   *
   * @private
   * @memberof Msal2Provider
   */
  private setStoredAccount() {
    this.clearStoredAccount();
    window[this.ms_config.cache.cacheLocation].setItem(
      this.homeAccountKey,
      this._publicClientApplication.getActiveAccount().homeAccountId
    );
  }

  /**
   * Get the stored account from storage
   *
   * @private
   * @return {*}
   * @memberof Msal2Provider
   */
  private getStoredAccount() {
    let homeId = null;

    homeId = window[this.ms_config.cache.cacheLocation].getItem(this.homeAccountKey);

    return this._publicClientApplication.getAccountByHomeId(homeId);
  }

  /**
   * Clears the stored account from storage
   *
   * @private
   * @memberof Msal2Provider
   */
  private clearStoredAccount() {
    window[this.ms_config.cache.cacheLocation].removeItem(this.homeAccountKey);
  }

  /**
   * Adds scopes that have already been requested to sessionstorage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof Msal2Provider
   */
  protected setRequestedScopes(scopes: string[]) {
    if (scopes) {
      sessionStorage.setItem(this.sessionStorageRequestedScopesKey, JSON.stringify(scopes));
    }
  }

  /**
   * Adds denied scopes to session storage
   *
   * @protected
   * @param {string[]} scopes
   * @memberof Msal2Provider
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
   * Gets denied scopes
   *
   * @protected
   * @return {*}
   * @memberof Msal2Provider
   */
  protected getDeniedScopes() {
    const scopesStr = sessionStorage.getItem(this.sessionStorageDeniedScopesKey);
    return scopesStr ? JSON.parse(scopesStr) : null;
  }

  /**
   * Checks if scopes were denied previously
   *
   * @protected
   * @param {string[]} scopes
   * @return {*}
   * @memberof Msal2Provider
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

  /**
   * Clears all requested scopes from session storage
   *
   * @protected
   * @memberof Msal2Provider
   */
  protected clearRequestedScopes() {
    sessionStorage.removeItem(this.sessionStorageRequestedScopesKey);
  }

  /**
   * Gets stored account if available, otherwise fetches the first account in the list of signed in accounts
   *
   * @private
   * @return {*}  {(AccountInfo | null)}
   * @memberof Msal2Provider
   */
  private getAccount(): AccountInfo | null {
    const account = this.getStoredAccount();
    if (account) {
      return account;
    } else if (this._publicClientApplication.getAllAccounts().length > 0) {
      return this._publicClientApplication.getAllAccounts()[0];
    }
    return null;
  }

  /**
   * Logs out user
   *
   * @memberof Msal2Provider
   */
  public async logout() {
    const logOutAccount = this._publicClientApplication.getActiveAccount();
    const logOutRequest: EndSessionRequest = {
      account: logOutAccount
    };
    this.clearStoredAccount();
    if (this._loginType == LoginType.Redirect) {
      this._publicClientApplication.logoutRedirect(logOutRequest);
    } else {
      this._publicClientApplication.logoutPopup({ ...logOutRequest });
    }
    this.setState(ProviderState.SignedOut);
  }

  /**
   * Returns access token for scopes
   *
   * @param {AuthenticationProviderOptions} [options]
   * @return {*}  {Promise<string>}
   * @memberof Msal2Provider
   */
  public async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest = {
      scopes: scopes,
      loginHint: this._loginHint,
      domainHint: this._domainHint
    };
    try {
      const silentRequest: SilentRequest = accessTokenRequest;
      const response = await this._publicClientApplication.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        if (this._loginType === LoginType.Redirect) {
          if (!this.areScopesDenied(scopes)) {
            this.setRequestedScopes(scopes);
            this._publicClientApplication.acquireTokenRedirect(accessTokenRequest);
          } else {
            throw e;
          }
        } else {
          try {
            const response = await this._publicClientApplication.acquireTokenPopup(accessTokenRequest);
            return response.accessToken;
          } catch (e) {
            throw e;
          }
        }
      } else {
        // if we don't know what the error is, just ask the user to sign in again
        this.setState(ProviderState.SignedOut);
      }
    }

    throw null;
  }
}
