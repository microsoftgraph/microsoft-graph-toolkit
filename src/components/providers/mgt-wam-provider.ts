/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement, property } from 'lit-element';
import { Providers } from '../../mgt-core';
import { WamProvider } from '../../providers/WamProvider';
/**
 * Authentication Library Provider for Web Account Manager (UWP apps)
 *
 * @export
 * @class MgtWamProvider
 * @extends {LitElement}
 */
@customElement('mgt-wam-provider')
export class MgtWamProvider extends LitElement {
  /**
   * String alphanumerical value relation to a specific user
   *
   * @type {string}
   * @memberof MgtWamProvider
   */
  @property({ attribute: 'client-id' }) public clientId: string;
  /**
   * Url used for login
   *
   * @type {string}
   * @memberof MgtWamProvider
   */
  @property({ attribute: 'authority' }) public authority?: string;
  /**
   * Invoked when the element is first updated and performs validation
   *
   * @param {*} changedProperties
   * @memberof MgtWamProvider
   */
  public firstUpdated(changedProperties) {
    this.validateAuthProps();
  }

  private validateAuthProps() {
    if (this.clientId !== undefined) {
      Providers.globalProvider = new WamProvider(this.clientId, this.authority);
    }
  }
}
