import { IProvider, LoginType, ProviderState, createFromProvider, Providers } from '@microsoft/mgt-element';
import {
  Configuration,
  LogLevel,
  PublicClientApplication,
  SilentRequest,
  PopupRequest,
  RedirectRequest,
  AuthenticationResult,
  AccountInfo
} from '@azure/msal-browser';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';

interface Msal2Config {
  clientId: string;
  loginType?: LoginType;
  scopes?: string[];
  loginHint?: string;
  redirectUri?: string;
  //TODO figure out why we need this twice
  options?: Configuration;
}
export class Msal2Provider extends IProvider {
  private _publicClientApplication: PublicClientApplication;
  private _loginType: LoginType;
  private _loginHint;
  private ms_config: Configuration;
  public scopes: string[];
  account: any;
  public constructor(config: Msal2Config) {
    super();

    this.initProvider(config);
  }

  private initProvider(config: Msal2Config) {
    if (config.clientId) {
      this.ms_config = {
        auth: {
          clientId: config.clientId,
          redirectUri: config.redirectUri
        },
        cache: {
          //TODO : COnfigure these
          cacheLocation: 'localStorage',
          storeAuthStateInCookie: false
        },
        system: {
          loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {},
            piiLoggingEnabled: false,
            logLevel: LogLevel.Warning
          }
        }
      };
      this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;
      this._loginHint = config.loginHint;
      this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
      this._publicClientApplication = new PublicClientApplication(this.ms_config);
      this._publicClientApplication
        .handleRedirectPromise()
        .then((tokenResponse: AuthenticationResult) => {
          if (tokenResponse !== null) {
            if (tokenResponse.idToken) {
              Providers.globalProvider.setState(ProviderState.SignedIn);
              this.handleResponse(tokenResponse);
            }
          }
        })
        .catch(console.error);
    } else {
      throw new Error('clientId must be provided');
    }

    //TODO : Can someone provide a PublicClientApplication?
    this.graph = createFromProvider(this);
  }

  //TODO: AuthenticationParameters
  public async login(): Promise<void> {
    const loginRequest: PopupRequest = {
      scopes: this.scopes,
      loginHint: this._loginHint,
      prompt: 'select_account'
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

  private getAccount(): AccountInfo | null {
    // need to call getAccount here?
    const currentAccounts = this._publicClientApplication.getAllAccounts();
    if (currentAccounts === null) {
      console.log('No accounts detected');
      return null;
    }

    if (currentAccounts.length > 1) {
      // Add choose account code here
      console.log('Multiple accounts detected, need to add choose account code.');
      return currentAccounts[0];
    } else if (currentAccounts.length === 1) {
      return currentAccounts[0];
    }

    return null;
  }

  public async logout() {
    this._publicClientApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  public async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    const scopes = options ? options.scopes || this.scopes : this.scopes;
    const accessTokenRequest = {
      loginHint: this._loginHint,
      scopes: scopes
    };
    try {
      //TODO: Figure out forerefresh
      const silentRequest: SilentRequest = { scopes: accessTokenRequest.scopes, forceRefresh: false };
      silentRequest.account = this.getAccount();
      const response = await this._publicClientApplication.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (e) {
      if (this.requiresInteraction(e)) {
        if (this._loginType === LoginType.Redirect) {
          //TODO : Check if user has denied scopes before
          console.log('acquiring token using redirect');
          this._publicClientApplication.acquireTokenRedirect(accessTokenRequest);
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
        throw e;
      }
    }

    throw null;
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
}
