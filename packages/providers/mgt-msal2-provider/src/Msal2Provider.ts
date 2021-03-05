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
  domain_hint?: string;
  redirectUri?: string;
  authority?: string;
  //TODO figure out why we need this twice
  options?: Configuration;
}
export class Msal2Provider extends IProvider {
  private _publicClientApplication: PublicClientApplication;
  private _loginType: LoginType;
  private _loginHint;
  private _domainHint;
  private ms_config: Configuration;
  public scopes: string[];
  account: any;

  // session storage
  private sessionStorageRequestedScopesKey = 'mgt-requested-scopes';
  private sessionStorageDeniedScopesKey = 'mgt-denied-scopes';
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
      //TODO: Add logger options
      this.ms_config.system.iframeHashTimeout = msalConfig.system.iframeHashTimeout || 10000;
      this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
      this._loginHint = typeof config.loginHint !== 'undefined' ? config.loginHint : null;
      this._domainHint = typeof config.domain_hint !== 'undefined' ? config.domain_hint : null;
      this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
      this._publicClientApplication = new PublicClientApplication(this.ms_config);
      this._publicClientApplication
        .handleRedirectPromise()
        .then(async (tokenResponse: AuthenticationResult) => {
          console.log('Inside redirect promise', tokenResponse);
          if (tokenResponse !== null) {
            if (tokenResponse.idToken) {
              this.clearRequestedScopes();
              Providers.globalProvider.setState(ProviderState.SignedIn);
              this.handleResponse(tokenResponse);
            }
          } else {
            await this.trySilentSignIn();
          }
        })
        .catch(async e => {
          console.log(e);
          console.log('Inside redirect error');
        });
    } else {
      throw new Error('clientId must be provided');
    }

    this.graph = createFromProvider(this);
  }

  public async trySilentSignIn() {
    this.account = this.getAccount();
    console.log('Inside silent before request', this.account);
    if (this.account) {
      const silentRequest = {
        scopes: this.scopes,
        sid: this.account.idTokenClaims.sid || null,
        loginHint: this._loginHint || this.account.idTokenClaims.preferred_username,
        domainHint: this._domainHint
      };
      console.log(silentRequest);
      this._publicClientApplication
        .ssoSilent(silentRequest)
        .then(() => {
          console.log('SILENT SUCCESSFUL!!!');
          if (this.account) {
            this.setState(ProviderState.SignedIn);
          } else {
            this.setState(ProviderState.SignedOut);
          }
        })
        .catch(async error => {
          console.log('Silent', error);
        });
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
      this.setState(response.account ? ProviderState.SignedIn : ProviderState.SignedOut);
    } else {
      //TODO: Multiple types of Requests
      const loginRedirectRequest: RedirectRequest = { ...loginRequest };
      this._publicClientApplication.loginRedirect(loginRedirectRequest);
    }
  }

  handleResponse(response: AuthenticationResult | null) {
    if (response !== null) {
      this.account = response.account;
    } else {
      this.account = this.getAccount();
    }
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
    // need to call getAccount here?
    const currentAccounts = this._publicClientApplication.getAllAccounts();
    if (currentAccounts === null) {
      console.log('No accounts detected');
      return null;
    }
    console.log(currentAccounts);
    if (currentAccounts.length > 1) {
      // Add choose account code here
      console.log('Multiple accounts detected, need to add choose account code.');
      if (this.account && typeof this.account.homeAccountId !== 'undefined') {
        console.log('Getting by home id');
        return this._publicClientApplication.getAccountByHomeId(this.account.homeAccountId);
      }
      return currentAccounts[0];
    } else if (currentAccounts.length === 1) {
      return currentAccounts[0];
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
      //TODO: Figure out forcerefresh
      const silentRequest: SilentRequest = accessTokenRequest;
      silentRequest.account = this.getAccount();
      const response = await this._publicClientApplication.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        if (this._loginType === LoginType.Redirect) {
          //TODO : Check if user has denied scopes before
          //TODO : Access token requests for redirect and popup are the same here
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
