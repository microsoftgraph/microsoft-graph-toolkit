import { IProvider, LoginType, ProviderState } from './IProvider';
import { Graph } from '../Graph';
import { UserAgentApplication } from 'msal/lib-es6';

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

  private _resolveToken;
  private _rejectToken;

  get provider() {
    return this._userAgentApplication;
  }

  scopes: string[];

  graph: Graph;

  constructor(config: MsalConfig) {
    super();
    this.initProvider(config);
  }


  private initProvider(config: MsalConfig) {
    console.log('initProvider');

    this.scopes =
      typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this._loginType =
      typeof config.loginType !== 'undefined'
        ? config.loginType
        : LoginType.Redirect;

    let callbackFunction = ((
      errorDesc: string,
      token: string,
      error: any,
      tokenType: any,
      state: any
    ) => {
      this.tokenReceivedCallback(errorDesc, token, error, tokenType, state);
    }).bind(this);

    if (config.userAgentApplication) {
      this._userAgentApplication = config.userAgentApplication;
    } else if (config.clientId) {
      let authority =
        typeof config.authority !== 'undefined' ? config.authority : null;
      let options =
        typeof config.options != 'undefined'
        ? config.options
        : { cacheLocation: 'localStorage' };

    this._userAgentApplication = new UserAgentApplication(
      config.clientId,
      authority,
      callbackFunction,
      options
    );
    } else {
      throw 'clientId or userAgentApplication must be provided';
    }
    
    this.graph = new Graph(this);

    this.tryGetIdTokenSilent();
  }

  async login(): Promise<void> {
    console.log('login');
    if (this._loginType === LoginType.Popup) {
      this._idToken = await this._userAgentApplication.loginPopup(this.scopes);
    } else {
      this._userAgentApplication.loginRedirect(this.scopes);
    }
  }

  async tryGetIdTokenSilent(): Promise<boolean> {
    console.log('tryGetIdTokenSilent');
    try {
      this._idToken = await this._userAgentApplication.acquireTokenSilent(
        [this._userAgentApplication.clientId],
        this._userAgentApplication.authority
      );
      if (this._idToken) {
        console.log('tryGetIdTokenSilent: got a token');
      }
      this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
      return this._idToken !== null;
    } catch (e) {
      console.log(e);
      this.setState(ProviderState.SignedOut);
      return false;
    }
  }

  private temp = 0;
  async getAccessToken(...scopes: string[]): Promise<string> {
    ++this.temp;
    let temp = this.temp;
    scopes = scopes || this.scopes;
    console.log('getaccesstoken' + ++temp + ': scopes' + scopes);
    let accessToken: string;
    try {
      accessToken = await this._userAgentApplication.acquireTokenSilent(
        scopes,
        this._userAgentApplication.authority
      );
      console.log('getaccesstoken' + temp + ': got token');
    } catch (e) {
      try {
        console.log('getaccesstoken' + temp + ': catch ' + e);
        // TODO - figure out for what error this logic is needed so we
        // don't prompt the user to login unnecessarily
        if (e.includes('multiple_matching_tokens_detected')) {
          console.log('getaccesstoken' + temp + ' ' + e);
          return null;
        }

        if (this._loginType == LoginType.Redirect) {
          this._userAgentApplication.acquireTokenRedirect(scopes);
          return new Promise((resolve, reject) => {
            this._resolveToken = resolve;
            this._rejectToken = reject;
          });
        } else {
          accessToken = await this._userAgentApplication.acquireTokenPopup(scopes);
        }
      } catch (e) {
        // TODO - figure out how to expose this during dev to make it easy for the dev to figure out
        // if error contains "'token' is not enabled", make sure to have implicit oAuth enabled in the AAD manifest
        console.log('getaccesstoken' + temp + 'catch2: ' + e);
        throw e;
      }
    }
    return accessToken;
  }

  async logout(): Promise<void> {
    this._userAgentApplication.logout();
    this.setState(ProviderState.SignedOut);
  }

  updateScopes(scopes: string[]) {
    this.scopes = scopes;
  }

  tokenReceivedCallback(
    errorDesc: string,
    token: string,
    error: any,
    tokenType: any,
    state: any
  ) {
    // debugger;
    console.log('tokenReceivedCallback ' + errorDesc + ' | ' + tokenType);
    if (this._userAgentApplication && window) {
      console.log(window.location.hash);
      console.log(
        'isCallback: ' + this._userAgentApplication.isCallback(window.location.hash)
      );
    }
    if (error) {
      console.log(error + ' ' + errorDesc);
      if (this._rejectToken) {
        this._rejectToken(errorDesc);
      }
    } else {
      if (tokenType == 'id_token') {
        this._idToken = token;
        this.setState(this._idToken ? ProviderState.SignedIn : ProviderState.SignedOut);
      } else {
        if (this._resolveToken) {
          this._resolveToken(token);
        }
      }
    }
  }
}
