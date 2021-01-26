import { IProvider, Providers, ProviderState, createFromProvider } from '@microsoft/mgt-element';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { ipcRenderer } from 'electron';

/**
 * ElectronProvider class to be instantiated in the renderer process.
 * Responsible for communicating with ElectronAuthenticator in the main process to acquire tokens
 *
 * @export
 * @class ElectronProvider
 * @extends {IProvider}
 */
export class ElectronProvider extends IProvider {
  constructor() {
    super();
    this.graph = createFromProvider(this);
    this.setupProvider();
  }

  /**
   * Sets up messaging between main and renderer to receive SignedIn/SignedOut state information
   *
   * @memberof ElectronProvider
   */
  setupProvider() {
    ipcRenderer.on('isloggedin', async (event, isLoggedIn) => {
      if (isLoggedIn) {
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } else {
        Providers.globalProvider.setState(ProviderState.SignedOut);
      }
    });
  }

  /**
   * Gets access token (called by MGT components)
   *
   * @param {AuthenticationProviderOptions} [options]
   * @return {*}  {Promise<string>}
   * @memberof ElectronProvider
   */
  async getAccessToken(options?: AuthenticationProviderOptions): Promise<string> {
    const token = await ipcRenderer.invoke('token', options);
    return token;
  }

  /**
   * Log in to set account information (called by mgt-login)
   *
   * @return {*}  {Promise<void>}
   * @memberof ElectronProvider
   */
  async login(): Promise<void> {
    Providers.globalProvider.setState(ProviderState.Loading);
    await ipcRenderer.invoke('login');
  }

  /**
   * Log out (called by mgt-login)
   *
   * @return {*}  {Promise<void>}
   * @memberof ElectronProvider
   */
  async logout(): Promise<void> {
    await ipcRenderer.invoke('logout');
  }
}
