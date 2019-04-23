import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { UserAgentApplication } from 'msal/lib-es6/UserAgentApplication';
import { IProvider, LoginType, ProviderState } from './IProvider';
import { Graph } from '../Graph';

export interface MsalConfig {
  userAgentApplication?: UserAgentApplication;
  clientId?: string;
  scopes?: string[];
  authority?: string;
  loginType?: LoginType;
  options?: any;
}

export class MsalProvider extends IProvider {
  private _loginType: LoginType;

  private _idToken: string;

  private _userAgentApplication: UserAgentApplication;

  get provider() {
    return this._userAgentApplication;
  }

  scopes: string[];

  constructor(config: MsalConfig) {
    super();
    this.initProvider(config);
  }

  private initProvider(config: MsalConfig) {
    this.scopes = typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType = typeof config.loginType !== 'undefined' ? config.loginType : LoginType.Redirect;

    let callbackFunction = ((errorDesc: string, token: string, error: any, tokenType: any, state: any) => {
      this.tokenReceivedCallback(errorDesc, token, error, tokenType, state);
    }).bind(this);

    if (config.userAgentApplication) {
      this._userAgentApplication = config.userAgentApplication;
    } else if (config.clientId) {
      let authority = typeof config.authority !== 'undefined' ? config.authority : null;
      let options =
        typeof config.options != 'undefined'
          ? config.options
          : { storeAuthStateInCookie: true, cacheLocation: 'localStorage' };

      this._userAgentApplication = new UserAgentApplication(config.clientId, authority, callbackFunction, options);
    } else {
      throw 'clientId or userAgentApplication must be provided';
    }

    this.graph = new Graph(this);

    this.tryGetIdTokenSilent();
  }

  async login(): Promise<void> {
    if (this._loginType === LoginType.Popup) {
      this._idToken = await this._userAgentApplication.loginPopup(this.scopes);
    } else {
      this._userAgentApplication.loginRedirect(this.scopes);
    }
  }

  async logout(): Promise<void> {
    this._userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  async tryGetIdTokenSilent(): Promise<boolean> {
    try {
      this._idToken = await this._userAgentApplication.acquireTokenSilent(
        [this._userAgentApplication.clientId],
        this._userAgentApplication.authority
      );
      if (this._idToken) {
      }
      this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
      return this._idToken !== null;
    } catch (e) {
      console.log(e);
      this.setState(ProviderState.SignedOut);
      return false;
    }
  }

  async getAccessToken(options: AuthenticationProviderOptions): Promise<string> {
    let scopes = options ? options.scopes || this.scopes : this.scopes;
    let accessToken: string;
    try {
      accessToken = await this._userAgentApplication.acquireTokenSilent(scopes, this._userAgentApplication.authority);
    } catch (e) {
      try {
        // TODO - figure out for what error this logic is needed so we
        // don't prompt the user to login unnecessarily
        if (e.includes('multiple_matching_tokens_detected')) {
          return null;
        }

        // AADSTS65001: The user or administrator has not consented to use the application
        // Need to send an interaction request
        if (e.includes('AADSTS65001')) {
          if (this._loginType == LoginType.Redirect) {
            // check if the user denied the scope before
            if (!this.areScopesDenied(scopes)) {
              this.setRequestedScopes(scopes);
              this._userAgentApplication.acquireTokenRedirect(scopes);
            }
          } else {
            accessToken = await this._userAgentApplication.acquireTokenPopup(scopes);
          }
        }
      } catch (e) {
        // TODO - figure out how to expose this during dev to make it easy for the dev to figure out
        // if error contains "'token' is not enabled", make sure to have implicit oAuth enabled in the AAD manifest
        console.log('getaccesstoken catch2 : ' + e);
        throw e;
      }
    }
    return accessToken;
  }

  updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  tokenReceivedCallback(errorDesc: string, token: string, error: any, tokenType: any, state: any) {
    if (error) {
      let requestedScopes = this.getRequestedScopes();
      if (requestedScopes) {
        this.addDeniedScopes(requestedScopes);
      }
    } else {
      if (tokenType == 'id_token') {
        this._idToken = token;
        this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
      } else {
      }
    }

    this.clearRequestedScopes();
  }

  // session storage
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
