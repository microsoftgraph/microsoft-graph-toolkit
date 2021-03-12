import { IProvider, LoginType, ProviderState, createFromProvider, Providers } from '@microsoft/mgt-element';
import {
  Configuration,
  LogLevel,
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
  private _loginType: LoginType;
  private _loginHint;
  private _domainHint;
  private _sid;
  private ms_config: Configuration;
  public scopes: string[];
  account: any;

  // session storage
  private sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
  private sessionStorageDeniedScopesKey = 'mgt-denied-scopes';
  private homeAccountKey = 'home-account-id';
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
          this.handleResponse(tokenResponse);
          if (tokenResponse === null) {
            this.trySilentSignIn();
          }
        });
      } catch (e) {
        console.log(e);
      }
    } else {
      throw new Error('clientId must be provided');
    }

    this.graph = createFromProvider(this);
  }

  public async trySilentSignIn() {
    //this.account = this.getAccount(); //TODO : Add else in case user provides sid- add to config
    console.log('Reached silent ');
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

  //TODO: AuthenticationParameters passed as argument
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

  handleResponse(response: AuthenticationResult | null) {
    console.log('All accounts', this._publicClientApplication.getAllAccounts());
    if (response !== null) {
      this.account = response.account;
      this.setState(ProviderState.SignedIn);
      this._publicClientApplication.setActiveAccount(this.account);
      this.setStoredAccount();
      console.log('Inside handleresponse', this.getStoredAccount());
    } else {
      this.setState(ProviderState.SignedOut);
    }
    this.clearRequestedScopes();
  }
  private setStoredAccount() {
    this.clearStoredAccount();
    localStorage.setItem(this.homeAccountKey, this._publicClientApplication.getActiveAccount().homeAccountId);
  }
  private getStoredAccount() {
    console.log('getstorecaccount');
    //TODO : Should I check if user has enabled localstorage? i.e.
    //if(this.ms_config.cache.cacheLocation === 'localstorage')
    const homeId = localStorage.getItem(this.homeAccountKey);
    return this._publicClientApplication.getAccountByHomeId(homeId);
  }

  private clearStoredAccount() {
    localStorage.removeItem(this.homeAccountKey);
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
    console.log('inside getaccount', this.getStoredAccount());
    if (account) {
      console.log('Account already exists inside getAccount', account);
      return account;
    }
    return null;
  }

  public async logout() {
    let account: AccountInfo | undefined;
    if (this.account) {
      account = this.account;
    }
    const logOutRequest: EndSessionRequest = {
      account
    };

    this._publicClientApplication.logout(logOutRequest);
    this.setState(ProviderState.SignedOut);
  }

  public async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    console.log('called', options ? options.scopes : 'Not from login');
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest = {
      scopes: scopes,
      loginHint: this._loginHint,
      domainHint: this._domainHint
    };
    try {
      const silentRequest: SilentRequest = accessTokenRequest;
      silentRequest.account = this.getAccount();
      console.log('Inside access token', silentRequest);
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
