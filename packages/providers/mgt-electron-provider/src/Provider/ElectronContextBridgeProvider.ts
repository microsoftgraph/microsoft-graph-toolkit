/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  GraphEndpoint,
  IProvider,
  MICROSOFT_GRAPH_DEFAULT_ENDPOINT,
  ProviderState,
  Providers,
  createFromProvider
} from '@microsoft/mgt-element';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
import type { IpcRendererEvent } from 'electron';

/**
 * The interface that describes the shape of the context bridge
 * that is used to communicate between the main and renderer processes.
 */
export interface IContextBridgeImpl {
  mgtAuthState: (callback: (event: IpcRendererEvent, authState: string) => void) => void;
  token: (options?: AuthenticationProviderOptions) => Promise<string>;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  approvedScopes: (callback: (event: IpcRendererEvent, approvedScopes: string[]) => void) => void;
}

/**
 * ElectronProvider class to be instantiated in the a preload script.
 * Responsible for communicating with ElectronAuthenticator in the main process to acquire tokens.
 *
 * This class uses the `contextBridge` to communicate with the main process,
 * which must be passed in as a constructor parameter.
 *
 * @export
 * @class ElectronProvider
 * @extends {IProvider}
 */
export class ElectronContextBridgeProvider extends IProvider {
  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtElectronContextBridgeProvider';
  }

  private readonly contextBridge: IContextBridgeImpl;

  constructor(contextBridge: IContextBridgeImpl, baseUrl: GraphEndpoint = MICROSOFT_GRAPH_DEFAULT_ENDPOINT) {
    super();
    this.baseURL = baseUrl;
    this.contextBridge = contextBridge;
    this.graph = createFromProvider(this);
    this.setupProvider();
  }

  /**
   * Sets up messaging between main and renderer to receive SignedIn/SignedOut state information
   *
   * @memberof ElectronProvider
   */
  setupProvider() {
    this.contextBridge.mgtAuthState((event, authState) => {
      if (authState === 'logged_in') {
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } else if (authState === 'logged_out') {
        Providers.globalProvider.setState(ProviderState.SignedOut);
      }
    });
    this.contextBridge.approvedScopes((_event, approvedScopes) => {
      Providers.globalProvider.approvedScopes = approvedScopes;
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
    const token = await this.contextBridge.token(options);
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
    await this.contextBridge.login();
  }

  /**
   * Log out (called by mgt-login)
   *
   * @return {*}  {Promise<void>}
   * @memberof ElectronProvider
   */
  async logout(): Promise<void> {
    await this.contextBridge.logout();
  }
}
