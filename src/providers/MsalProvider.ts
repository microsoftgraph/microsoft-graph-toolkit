import { IProvider, LoginType, ProviderState } from './IProvider';
import { Graph } from '../Graph';
import { UserAgentApplication } from 'msal/lib-es6';

export interface MsalConfig {
  clientId: string;
  scopes?: string[];
  authority?: string;
  loginType?: LoginType;
  options?: any;
}

export class MsalProvider extends IProvider {
  private _loginType: LoginType;
  private _clientId: string;

  private _idToken: string;

  private _provider: UserAgentApplication;

  private _resolveToken;
  private _rejectToken;

  get provider() {
    return this._provider;
  }

  scopes: string[];
  authority: string;

  graph: Graph;

  constructor(config: MsalConfig) {
    super();
    if (!config.clientId) {
      throw 'ClientID must be a valid string';
    }

    this.initProvider(config);
  }

  private initProvider(config: MsalConfig) {
    console.log('initProvider');
    this._clientId = config.clientId;
    this.scopes =
      typeof config.scopes !== 'undefined' ? config.scopes : ['user.read'];
    this.authority =
      typeof config.authority !== 'undefined' ? config.authority : null;
    let options =
      typeof config.options != 'undefined'
        ? config.options
        : { cacheLocation: 'localStorage' };
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

    this._provider = new UserAgentApplication(
      this._clientId,
      this.authority,
      callbackFunction,
      options
    );
    this.graph = new Graph(this);

    this.tryGetIdTokenSilent();
  }

  async login(): Promise<void> {
    console.log('login');
    if (this._loginType == LoginType.Popup) {
      this._idToken = await this.provider.loginPopup(this.scopes);
    } else {
      this.provider.loginRedirect(this.scopes);
    }
  }

  async tryGetIdTokenSilent(): Promise<boolean> {
    console.log('tryGetIdTokenSilent');
    try {
      this._idToken = await this.provider.acquireTokenSilent(
        [this._clientId],
        this.authority
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
      accessToken = await this.provider.acquireTokenSilent(
        scopes,
        this.authority
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
          this.provider.acquireTokenRedirect(scopes);
          return new Promise((resolve, reject) => {
            this._resolveToken = resolve;
            this._rejectToken = reject;
          });
        } else {
          accessToken = await this.provider.acquireTokenPopup(scopes);
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
    this.provider.logout();
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
    if (this._provider && window) {
      console.log(window.location.hash);
      console.log(
        'isCallback: ' + this._provider.isCallback(window.location.hash)
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
