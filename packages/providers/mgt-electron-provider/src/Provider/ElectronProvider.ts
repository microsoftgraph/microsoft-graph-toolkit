/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  IProvider,
  Providers,
  ProviderState,
  createFromProvider,
  GraphEndpoint,
  MICROSOFT_GRAPH_DEFAULT_ENDPOINT
} from '@microsoft/mgt-element';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
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
  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtElectronProvider';
  }

  constructor(baseUrl: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT) {
    super();
    this.baseURL = baseUrl;
    this.graph = createFromProvider(this);
    this.setupProvider();
  }

  /**
   * Sets up messaging between main and renderer to receive SignedIn/SignedOut state information
   *
   * @memberof ElectronProvider
   */
  setupProvider() {
    ipcRenderer.on('mgtAuthState', (event, authState) => {
      if (authState === 'logged_in') {
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } else if (authState === 'logged_out') {
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
    const token = (await ipcRenderer.invoke('token', options)) as string;
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
