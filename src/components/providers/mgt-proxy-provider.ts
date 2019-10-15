/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement, property } from 'lit-element';
import { Providers } from '../../Providers';
import { ProxyProvider } from '../../providers/ProxyProvider';
/**
 * Authentication Library Provider for Proxy Auth thru an api server
 *
 * @export
 * @class MgtProxyProvider
 * @extends {LitElement}
 */
@customElement('mgt-proxy-provider')
export class MgtProxyProvider extends LitElement {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @type {string}
   * @memberof MgtProxyProvider
   */
  @property({ attribute: 'graph-proxy-url' }) public graphProxyUrl: string;
  /**
   * Invoked when the element is first updated and performs validation
   *
   * @param {*} changedProperties
   * @memberof MgtProxyProvider
   */
  public firstUpdated(changedProperties) {
    this.validateAuthProps();
  }

  private validateAuthProps() {
    if (this.graphProxyUrl !== undefined) {
      Providers.globalProvider = new ProxyProvider(this.graphProxyUrl);
    }
  }
}
