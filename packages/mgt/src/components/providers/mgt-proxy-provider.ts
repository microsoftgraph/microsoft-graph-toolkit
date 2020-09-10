/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { Providers } from '@microsoft/mgt-element';
import { ProxyProvider } from '../../providers/ProxyProvider';
import { MgtBaseProvider } from './baseProvider';

/**
 * Authentication component for ProxyProvider
 *
 * @export
 * @class MgtProxyProvider
 * @extends {LitElement}
 */
@customElement('mgt-proxy-provider')
export class MgtProxyProvider extends MgtBaseProvider {
  /**
   * The base url to the proxy api
   *
   * @type {string}
   * @memberof MgtProxyProvider
   */
  @property({ attribute: 'graph-proxy-url' }) public graphProxyUrl: string;

  /**
   * Gets whether this provider can be used in this environment
   *
   * @readonly
   * @memberof MgtMsalProvider
   */
  public get isAvailable() {
    return true;
  }

  /**
   * method called to initialize the provider. Each derived class should provide their own implementation.
   *
   * @protected
   * @memberof MgtProxyProvider
   */
  protected initializeProvider() {
    if (this.graphProxyUrl !== undefined) {
      this.provider = new ProxyProvider(this.graphProxyUrl);
      Providers.globalProvider = this.provider;
    }
  }
}
