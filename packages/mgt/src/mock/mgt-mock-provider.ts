/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property } from 'lit-element';
import { MgtBaseProvider } from '../components/providers/baseProvider';
import { Providers } from '@microsoft/mgt-element';
import { MockProvider } from './MockProvider';
/**
 * Sets global provider to a mock Provider
 *
 * @export
 * @class MgtMockProvider
 * @extends {LitElement}
 */
@customElement('mgt-mock-provider')
export class MgtMockProvider extends MgtBaseProvider {
  /**
   * A property to allow the developer to start the sample logged out if they desired.
   *
   * @memberof MgtMockProvider
   */
  @property({
    attribute: 'signed-out',
    type: Boolean
  })
  public signedOut;

  /**
   * method called to initialize the provider. Each derived class should provide
   * their own implementation
   *
   * @protected
   * @memberof MgtBaseProvider
   */
  protected initializeProvider() {
    Providers.globalProvider = new MockProvider(!this.signedOut);
  }
}
