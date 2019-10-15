/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Graph } from '../Graph';
import { IProvider, ProviderState } from '../providers/IProvider';
/**
 * Proxy Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class ProxyProvider
 * @extends {IProvider}
 */
export class ProxyProvider extends IProvider {
  // tslint:disable-next-line: completed-docs
  public provider: any;

  /**
   * new instance of proxy graph provider
   *
   * @memberof ProxyProvider
   */
  public graph: Graph;
  constructor(graphProxyUrl: string) {
    super();
    this.graph = new ProxyGraph(graphProxyUrl, this);
    this.graph.getMe().then(
      user => {
        if (user != null) {
          this.setState(ProviderState.SignedIn);
        } else {
          this.setState(ProviderState.SignedOut);
        }
      },
      err => {
        this.setState(ProviderState.SignedOut);
      }
    );
  }

  /**
   * sets Provider state to SignedIn
   *
   * @returns {Promise<void>}
   * @memberof ProxyProvider
   */
  public async login(): Promise<void> {
    this.setState(ProviderState.SignedIn);
  }
  /**
   * sets Provider state to signed out
   *
   * @returns {Promise<void>}
   * @memberof ProxyProvider
   */
  public async logout(): Promise<void> {
    this.setState(ProviderState.Loading);
    await new Promise(resolve => setTimeout(resolve, 3000));
    this.setState(ProviderState.SignedOut);
  }
  /**
   * Promise returning token from graph.microsoft.com
   *
   * @returns {Promise<string>}
   * @memberof ProxyProvider
   */
  public getAccessToken(): Promise<string> {
    return Promise.resolve('{token:https://graph.microsoft.com/}');
  }
}

/**
 * ProxyGraph Instance
 *
 * @export
 * @class ProxyGraph
 * @extends {Graph}
 */
// tslint:disable-next-line: max-classes-per-file
export class ProxyGraph extends Graph {
  private readonly baseUrl: string;

  constructor(baseUrl: string, provider: ProxyProvider) {
    super(null);

    this.baseUrl = baseUrl;

    this.client = Client.initWithMiddleware({
      authProvider: provider,
      baseUrl: this.baseUrl
    });
  }
}
