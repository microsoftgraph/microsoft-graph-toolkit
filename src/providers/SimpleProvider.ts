import { IProvider } from './IProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { Graph } from '../Graph';

export class SimpleProvider extends IProvider {
  private _getAccessTokenHandler: (scopes: string[]) => Promise<string>;
  private _loginHandler: () => Promise<void>;
  private _logoutHandler: () => Promise<void>;

  constructor(
    getAccessTokenHandler: (scopes: string[]) => Promise<string>,
    loginHandler?: () => Promise<void>,
    logoutHandler?: () => Promise<void>
  ) {
    super();

    this._getAccessTokenHandler = getAccessTokenHandler;
    this._loginHandler = loginHandler;
    this._logoutHandler = logoutHandler;

    this.graph = new Graph(this);
  }

  getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    return this._getAccessTokenHandler(options.scopes);
  }

  login(): Promise<void> {
    return this._loginHandler();
  }

  logout(): Promise<void> {
    return this._logoutHandler();
  }
}
