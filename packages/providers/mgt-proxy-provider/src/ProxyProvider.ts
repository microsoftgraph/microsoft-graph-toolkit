/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider, ProviderState, Graph } from '@microsoft/mgt-element';
import { ProxyGraph } from './ProxyGraph';

/**
 * Proxy Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class ProxyProvider
 * @extends {IProvider}
 */
export class ProxyProvider extends IProvider {
  /**
   * new instance of proxy graph provider
   *
   * @memberof ProxyProvider
   */
  public graph: Graph;

  /**
   * Name used for analytics
   *
   * @readonly
   * @memberof IProvider
   */
  public get name() {
    return 'MgtProxyProvider';
  }

  constructor(graphProxyUrl: string, getCustomHeaders: () => Promise<object> = null) {
    super();
    this.graph = new ProxyGraph(graphProxyUrl, getCustomHeaders);

    this.graph
      .api('me')
      .get()
      .then(
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
   * Promise returning token
   *
   * @returns {Promise<string>}
   * @memberof ProxyProvider
   */
  public getAccessToken(): Promise<string> {
    return null;
  }
}
