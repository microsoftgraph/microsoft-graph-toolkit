/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  IProvider,
  LoginType,
  ProviderState,
  createFromProvider,
  IProviderAccount,
  GraphEndpoint
} from '@microsoft/mgt-element';
import {
  Configuration,
  PublicClientApplication,
  SilentRequest,
  PopupRequest,
  RedirectRequest,
  AccountInfo,
  EndSessionRequest,
  InteractionRequiredAuthError,
  SsoSilentRequest,
  EventMessage,
  AuthenticationResult,
  BrowserAuthError
} from '@azure/msal-browser';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';

/**
 * base config for MSAL 2.0 authentication
 *
 * @export
 * @interface Msal2ConfigBase
 */
interface Msal2ConfigBase {
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
   * Other options
   *
   * @type {Configuration}
   * @memberof Msal2Config
   */
  options?: Configuration;

  /**
   * List of scopes required
   *
   * @type {string[]}
   * @memberof Msal2ConfigBase
   */
  scopes?: string[];
  /**
   * loginType if login uses popup
   *
   * @type {LoginType}
   * @memberof Msal2ConfigBase
   */
  loginType?: LoginType;
  /**
   * login hint value
   *
   * @type {string}
   * @memberof Msal2ConfigBase
   */
  loginHint?: string;
  /**
   * Domain hint value
   *
   * @type {string}
   * @memberof Msal2ConfigBase
   */
  domainHint?: string;

  /**
   * Prompt type
   *
   * @type {string}
   * @memberof Msal2ConfigBase
   */
  prompt?: PromptType;

  /**
   * Session ID
   *
   * @type {string}
   * @memberof Msal2Config
   */
  sid?: string;

  /**
   * Specifies if incremental consent is disabled
   *
   * @type {boolean}
   * @memberof Msal2ConfigBase
   */
  isIncrementalConsentDisabled?: boolean;

  /**
   * Disable multi account functionality
   *
   * @type {boolean}
   * @memberof Msal2Config
   */
  isMultiAccountDisabled?: boolean;
}

/**
 * Config for MSAL2.0 Authentication
 *
 * @export
 * @interface Msal2Config
 */
export interface Msal2Config extends Msal2ConfigBase {
  /**
   * Client ID of app registration
   *
   * @type {boolean}
   * @memberof Msal2Config
   */
  clientId: string;

  /**
   * Disable multi account functionality
   *
   * @type {boolean}
   * @memberof Msal2Config
   */
  isMultiAccountEnabled?: boolean;

  /**
   * The base URL for the graph client
   */
  baseURL?: GraphEndpoint;

  /**
   * CustomHosts
   *
   * @type {string[]}
   * @memberof Msal2Config
   */
  customHosts?: string[];
}

/**
 * Config for MSAL 2.0 Authentication where a PublicClientApplication already exists
 *
 * @export
 * @interface Msal2PublicClientApplicationConfig
 */
export interface Msal2PublicClientApplicationConfig extends Msal2ConfigBase {
  /**
   * Existing PublicClientApplication instance to use
   *
   * @type {PublicClientApplication}
   * @memberof Msal2PublicClientApplicationConfig
   */
  publicClientApplication: PublicClientApplication;
}

/**
 * Prompt type enum
 *
 * @export
 * @enum {string}
 */
export enum PromptType {
  SELECT_ACCOUNT = 'select_account',
  LOGIN = 'login',
  CONSENT = 'consent'
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
  private _loginHint: string;

  /**
   * Domain hint if provided
   *
   * @private
   * @memberof Msal2Provider
   */
  private _domainHint: string;

  /**
   * Prompt type
   *
   * @private
   * @type {string}
   * @memberof Msal2Provider
   */
  private _prompt: string;

  /**
   * Session ID, if provided
   *
   * @private
   * @memberof Msal2Provider
   */
  private _sid: string;

  // /**
  //  * Specifies if incremental consent is disabled
  //  *
  //  * @type {boolean}
  //  * @memberof Msal2ConfigBase
  //  */
  // private _isIncrementalConsentDisabled: boolean = false;

  /**
   * Configuration settings for authentication
   *
   * @private
   * @type {Configuration}
   * @memberof Msal2Provider
   */
  // eslint-disable-next-line @typescript-eslint/naming-convention
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
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtMsal2Provider';
  }

  /**
   * List of scopes
   *
   * @type {string[]}
   * @memberof Msal2Provider
   */
  public scopes: string[];

  /**
   * Enables multi account functionality if true, disables if false
   *
   * @private
   * @type {boolean}
   * @memberof Msal2Provider
   */
  public isMultipleAccountEnabled = true;

  /**
   * Indicates if multi account functionality is disabled
   *
   * @protected
   * @type {boolean}
   * @memberof Msal2Provider
   */
  protected get isMultiAccountDisabled(): boolean {
    return !this.isMultipleAccountEnabled;
  }

  /**
   * Disables or enables multi account functionality
   * Uses isMultipleAccountEnabled as the backing property
   * Property provided to ensure adherence to the IProvider interface
   *
   * @protected
   * @memberof Msal2Provider
   */
  protected set isMultiAccountDisabled(value: boolean) {
    this.isMultipleAccountEnabled = !value;
  }

  /**
   * Specifies if Multi account functionality is supported by the provider and enabled.
   *
   * @readonly
   * @type {boolean}
   * @memberof IProvider
   */
  public get isMultiAccountSupportedAndEnabled(): boolean {
    return this.isMultipleAccountEnabled;
  }

  private get sessionStorageRequestedScopesKey() {
    return 'mgt-requested-scopes';
  }
  private get sessionStorageDeniedScopesKey() {
    return 'mgt-denied-scopes';
  }
  private get homeAccountKey() {
    return '275f3731-e4a4-468a-bf9c-baca24b31e26';
  }

  public constructor(config: Msal2Config | Msal2PublicClientApplicationConfig) {
    super();
    void this.initProvider(config);
  }

  /**
   * Initialize provider with configuration details
   *
   * @private
   * @param {Msal2Config} config
   * @memberof Msal2Provider
   */
  private async initProvider(config: Msal2Config | Msal2PublicClientApplicationConfig) {
    const msalConfig: Configuration = config.options || { auth: { clientId: '' } };
    this.ms_config = msalConfig;
    this.ms_config.cache = msalConfig.cache || {};
    this.ms_config.cache.cacheLocation = msalConfig.cache.cacheLocation || 'localStorage';
    if (
      typeof this.ms_config.cache.storeAuthStateInCookie === 'undefined' ||
      this.ms_config.cache.storeAuthStateInCookie === null
    ) {
      this.ms_config.cache.storeAuthStateInCookie = true;
    }

    this.ms_config.system = msalConfig.system || {};
    this.ms_config.system.iframeHashTimeout = msalConfig.system.iframeHashTimeout || 10000;

    if (config.authority) {
      this.ms_config.auth.authority = config.authority;
    }
    if (config.redirectUri) {
      this.ms_config.auth.redirectUri = config.redirectUri;
    }

    if ('clientId' in config) {
      if (config.clientId) {
        this.ms_config.auth.clientId = config.clientId;
        this._publicClientApplication = new PublicClientApplication(this.ms_config);
      } else {
        throw new Error('clientId must be provided');
      }
    } else if ('publicClientApplication' in config) {
      if (config.publicClientApplication) {
        this._publicClientApplication = config.publicClientApplication;
      } else {
        throw new Error('publicClientApplication must be provided');
      }
    } else {
      throw new Error('either clientId or publicClientApplication must be provided');
    }
    this._publicClientApplication.addEventCallback(this.handleMsalEvent);
    await this._publicClientApplication.initialize();

    this.ms_config.system = msalConfig.system || {};
    this.ms_config.system.iframeHashTimeout = msalConfig.system.iframeHashTimeout || 10000;
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
    this._loginHint = typeof config.loginHint !== 'undefined' ? config.loginHint : null;
    this._sid = typeof config.sid !== 'undefined' ? config.sid : null;
    this.isIncrementalConsentDisabled =
      typeof config.isIncrementalConsentDisabled !== 'undefined' ? config.isIncrementalConsentDisabled : false;
    this._domainHint = typeof config.domainHint !== 'undefined' ? config.domainHint : null;
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._prompt = typeof config.prompt !== 'undefined' ? config.prompt : PromptType.SELECT_ACCOUNT;

    const msal2config = config as Msal2Config;
    this.isMultipleAccountEnabled =
      typeof msal2config.isMultiAccountEnabled !== 'undefined' ? msal2config.isMultiAccountEnabled : true;
    this.baseURL = typeof msal2config.baseURL !== 'undefined' ? msal2config.baseURL : this.baseURL;
    this.customHosts = msal2config.customHosts;

    this.graph = createFromProvider(this);
    try {
      const tokenResponse = await this._publicClientApplication.handleRedirectPromise();
      if (tokenResponse !== null) {
        this.handleResponse(tokenResponse?.account);
      } else {
        await this.trySilentSignIn();
      }
    } catch (e) {
      console.error('🦒: Problem attempting to sign in', e);
      throw e;
    }
  }

  private readonly handleMsalEvent = (message: EventMessage): void => {
    if (message.eventType === 'msal:acquireTokenSuccess' && 'scopes' in message.payload) {
      this.approvedScopes = message.payload.scopes;
    }
  };

  /**
   * Attempts to sign in user silently
   *
   * @memberof Msal2Provider
   */
  public async trySilentSignIn() {
    const silentRequest: SsoSilentRequest = {
      scopes: this.scopes,
      domainHint: this._domainHint
    };
    if (this._sid || this._loginHint) {
      silentRequest.sid = this._sid;
      silentRequest.loginHint = this._loginHint;
      try {
        this.setState(ProviderState.Loading);
        const response = await this._publicClientApplication.ssoSilent(silentRequest);
        if (response) {
          this.handleResponse(response?.account);
        }
      } catch (e) {
        this.setState(ProviderState.SignedOut);
      }
    } else {
      const account: AccountInfo = this.getAccount();
      if (account) {
        if (await this.getAccessToken(null)) {
          this.handleResponse(account);
          return;
        }
      }
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
      prompt: this._prompt,
      domainHint: this._domainHint
    };
    if (this._loginType === LoginType.Popup) {
      try {
        const response = await this._publicClientApplication.loginPopup(loginRequest);
        this.handleResponse(response?.account);
      } catch (error) {
        switch (true) {
          case error instanceof BrowserAuthError && error.errorCode === 'user_cancelled':
            console.warn('🦒: User cancelled the login flow.');
            this.setState(ProviderState.SignedOut);
            break;
          case error instanceof BrowserAuthError && error.errorCode === 'interaction_in_progress':
            console.warn('🦒: Login already in progess. Close the popup to login again.');
            this.setState(ProviderState.SignedOut);
            break;
          default:
            console.error('🦒: Error occurred during login:', error);
            throw error;
        }
      }
    } else {
      const loginRedirectRequest: RedirectRequest = { ...loginRequest };
      await this._publicClientApplication.loginRedirect(loginRedirectRequest);
    }
  }

  /**
   * Get all signed in accounts
   *
   * @return {*}
   * @memberof Msal2Provider
   */
  public getAllAccounts() {
    const usernames: IProviderAccount[] = [];
    this._publicClientApplication.getAllAccounts().forEach((account: AccountInfo) => {
      usernames.push({ name: account.name, mail: account.username, id: account.homeAccountId } as IProviderAccount);
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
    const accountToSet = this._publicClientApplication.getAccountByHomeId(user.id);
    const activeAccount = this._publicClientApplication.getActiveAccount();
    const storedAccount = this.getStoredAccount();
    // exit early if the account is already active and stored
    if (
      storedAccount &&
      accountToSet &&
      activeAccount &&
      storedAccount.homeAccountId === accountToSet.homeAccountId &&
      activeAccount.homeAccountId === accountToSet.homeAccountId
    ) {
      return;
    }
    this._publicClientApplication.setActiveAccount(accountToSet);
    this.setStoredAccount();
    super.setActiveAccount(user);
  }

  /**
   * Gets active account
   *
   * @return {*}
   * @memberof Msal2Provider
   */
  public getActiveAccount() {
    const account = this._publicClientApplication.getActiveAccount();

    if (account) {
      return {
        name: account.name,
        mail: account.username,
        id: account.homeAccountId,
        tenantId: account.tenantId
      } as IProviderAccount;
    }
    return undefined;
  }

  /**
   * Once a succesful login occurs, set the active account and store it
   *
   * @param {(AuthenticationResult | null)} account
   * @memberof Msal2Provider
   */
  handleResponse(account: AccountInfo) {
    if (account !== null) {
      this.setActiveAccount({
        name: account.name,
        id: account.homeAccountId,
        mail: account.username
      } as IProviderAccount);
      this.setState(ProviderState.SignedIn);
    } else {
      this.setState(ProviderState.SignedOut);
    }
    this.clearRequestedScopes();
  }

  private storage() {
    switch (this.ms_config.cache.cacheLocation) {
      case 'localStorage':
        return window.localStorage;
      case 'sessionStorage':
      default:
        return window.sessionStorage;
    }
  }

  /**
   * Store the currently signed in account in storage
   *
   * @private
   * @memberof Msal2Provider
   */
  private setStoredAccount() {
    this.clearStoredAccount();
    this.storage().setItem(this.homeAccountKey, this._publicClientApplication.getActiveAccount().homeAccountId);
  }

  /**
   * Get the stored account from storage
   *
   * @private
   * @return {*}
   * @memberof Msal2Provider
   */
  private getStoredAccount() {
    const homeId = this.storage().getItem(this.homeAccountKey);

    return this._publicClientApplication.getAccountByHomeId(homeId);
  }

  /**
   * Clears the stored account from storage
   *
   * @private
   * @memberof Msal2Provider
   */
  private clearStoredAccount() {
    this.storage().removeItem(this.homeAccountKey);
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
    return scopesStr ? (JSON.parse(scopesStr) as string[]) : null;
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
  protected getAccount(): AccountInfo | null {
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
    if (this._loginType === LoginType.Redirect) {
      this.setState(ProviderState.SignedOut);
      await this._publicClientApplication.logoutRedirect(logOutRequest);
    } else {
      await this._publicClientApplication.logoutPopup({ ...logOutRequest });
      if (this._publicClientApplication.getAllAccounts.length === 1 || !this.isMultipleAccountEnabled) {
        this.setState(ProviderState.SignedOut);
      } else {
        await this.trySilentSignIn();
      }
    }
  }

  /**
   * Returns access token for scopes
   *
   * @param {AuthenticationProviderOptions} [options]
   * @return {*}  {Promise<string>}
   * @memberof Msal2Provider
   */
  public async getAccessToken(opts?: AuthenticationProviderOptions): Promise<string> {
    const scopes = opts?.scopes?.length ? opts.scopes : this.scopes;
    const accessTokenRequest: SilentRequest = {
      scopes,
      account: this.getAccount()
    };
    try {
      const silentRequest: SilentRequest = accessTokenRequest;
      const response = await this._publicClientApplication.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        if (this.isIncrementalConsentDisabled) {
          return null;
        }
        if (this._loginType === LoginType.Redirect) {
          if (!this.areScopesDenied(scopes)) {
            this.setRequestedScopes(scopes);
            await this._publicClientApplication.acquireTokenRedirect(accessTokenRequest);
          } else {
            throw e;
          }
        } else {
          try {
            const response = await this._publicClientApplication.acquireTokenPopup(accessTokenRequest);
            return response.accessToken;
          } catch (popUpErr) {
            console.error('🦒: problem with pop-up sign in', popUpErr);
            throw popUpErr;
          }
        }
      } else {
        // if we don't know what the error is, just ask the user to sign in again
        this.setState(ProviderState.SignedOut);
      }
    }
    // TODO: work out if we can throw something more meaningful here
    // eslint-disable-next-line no-throw-literal
    throw null;
  }
}
