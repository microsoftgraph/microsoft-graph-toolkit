import { IProvider, LoginType, ProviderState, createFromProvider, Providers } from '@microsoft/mgt-element';
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
import { isRegExp } from 'util';

export interface Msal2Config {
  clientId: string;
  loginType?: LoginType;
  scopes?: string[];
  loginHint?: string;
  sid?: string;
  domain_hint?: string;
  redirectUri?: string;
  authority?: string;
  options?: Configuration;
}
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

  public scopes: string[];
  account: any;

  // session storage

  //TODO: Bugs - switching accounts doesn't work with caching
  //TODO: Adding new account does not work with loginpopup
  private sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
  private sessionStorageDeniedScopesKey = 'mgt-denied-scopes';
  private homeAccountKey = '275f3731-e4a4-468a-bf9c-baca24b31e26';
  public constructor(config: Msal2Config) {
    super();
    this.initProvider(config);
  }

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
      try {
        this._publicClientApplication.handleRedirectPromise().then((tokenResponse: AuthenticationResult | null) => {
          console.log('Redirect');
          //Work on this later
          this.handleResponse(tokenResponse);
          if (tokenResponse === null) {
            this.trySilentSignIn();
          }
        });
      } catch (e) {
        console.log('redirect', e);
      }
    } else {
      throw new Error('clientId must be provided');
    }

    this.graph = createFromProvider(this);
  }

  public async trySilentSignIn() {
    let silentRequest: any = {
      scopes: this.scopes,
      domainHint: this._domainHint
    };
    if (this._sid || this._loginHint) {
      console.log('trying silent with sid & loginhint', this._sid, this._loginHint);
      silentRequest.sid = this._sid;
      silentRequest.loginHint = this._loginHint;
    } else {
      this.account = this.getAccount();
      if (this.account) {
        console.log('trying silent with account', this.account);
        (silentRequest.sid = this.account.idTokenClaims.sid || null),
          (silentRequest.loginHint = this.account.idTokenClaims.preferred_username);
      }
    }
    if (silentRequest.sid) {
      try {
        this.setState(ProviderState.Loading);
        const response = await this._publicClientApplication.ssoSilent(silentRequest);
        if (response) {
          console.log('SILENT SUCCESSFUL!!!');
          this.handleResponse(response);
        }
      } catch (e) {
        console.log('Silent', e);
        this.setState(ProviderState.SignedOut);
      }
    }
  }

  public async login(): Promise<void> {
    const loginRequest: PopupRequest = {
      scopes: this.scopes,
      loginHint: this._loginHint,
      prompt: 'select_account',
      domainHint: this._domainHint
    };
    this._publicClientApplication.setActiveAccount(null); //TODO: This is a temporary fix
    if (this._loginType == LoginType.Popup) {
      const response = await this._publicClientApplication.loginPopup(loginRequest);
      this.handleResponse(response);
      this.fireActiveAccountChanged();
    } else {
      const loginRedirectRequest: RedirectRequest = { ...loginRequest };
      this._publicClientApplication.loginRedirect(loginRedirectRequest);
    }
  }

  public getAllAccounts() {
    let usernames = [];
    this._publicClientApplication.getAllAccounts().forEach((account: AccountInfo) => {
      usernames.push({ username: account.username, id: account.homeAccountId });
    });
    return usernames;
  }

  public switchAccount(user: any) {
    this._publicClientApplication.setActiveAccount(this._publicClientApplication.getAccountByHomeId(user.id));
    this.setStoredAccount();
    console.log('New getStoredAccount', this.getStoredAccount());
    this.fireActiveAccountChanged();
  }

  handleResponse(response: AuthenticationResult | null) {
    if (response !== null) {
      this.account = response.account;
      this._publicClientApplication.setActiveAccount(this.account);
      this.setStoredAccount();
      this.setState(ProviderState.SignedIn);
      console.log(
        'All accounts',
        this._publicClientApplication.getAllAccounts(),
        'Active account',
        this._publicClientApplication.getActiveAccount()
      );
    } else {
      this.setState(ProviderState.SignedOut);
    }
    this.clearRequestedScopes();
  }

  private setStoredAccount() {
    this.clearStoredAccount();
    window[this.ms_config.cache.cacheLocation].setItem(
      this.homeAccountKey,
      this._publicClientApplication.getActiveAccount().homeAccountId
    );
  }
  private getStoredAccount() {
    let homeId = null;

    homeId = window[this.ms_config.cache.cacheLocation].getItem(this.homeAccountKey);

    return this._publicClientApplication.getAccountByHomeId(homeId);
  }

  private clearStoredAccount() {
    window[this.ms_config.cache.cacheLocation].removeItem(this.homeAccountKey);
  }

  protected setRequestedScopes(scopes: string[]) {
    if (scopes) {
      sessionStorage.setItem(this.sessionStorageRequestedScopesKey, JSON.stringify(scopes));
    }
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
      sessionStorage.setItem(this.sessionStorageDeniedScopesKey, JSON.stringify(deniedScopes));
    }
  }

  protected getDeniedScopes() {
    const scopesStr = sessionStorage.getItem(this.sessionStorageDeniedScopesKey);
    return scopesStr ? JSON.parse(scopesStr) : null;
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

  protected clearRequestedScopes() {
    sessionStorage.removeItem(this.sessionStorageRequestedScopesKey);
  }

  private getAccount(): AccountInfo | null {
    const account = this.getStoredAccount();
    console.log('Stored Account', this.getStoredAccount());
    if (account) {
      return account;
    } else if (this._publicClientApplication.getAllAccounts().length > 0) {
      return this._publicClientApplication.getAllAccounts()[0];
    }
    return null;
  }

  public async logout() {
    const logOutAccount = this._publicClientApplication.getActiveAccount();
    const logOutRequest: EndSessionRequest = {
      account: logOutAccount
    };
    this.clearStoredAccount();
    console.log(logOutRequest.account);
    this._publicClientApplication.logout(logOutRequest);
    this.setState(ProviderState.SignedOut);
  }

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
        console.error(e);
      }
    }

    throw null;
  }
}
