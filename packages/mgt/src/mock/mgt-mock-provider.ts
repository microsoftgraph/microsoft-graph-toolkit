/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property } from 'lit-element';
import { MgtBaseProvider } from '../components/providers/baseProvider';
import { Providers } from '../Providers';
import { registeredComponent } from '../utils/ComponentRegistry';
import { MockProvider } from './MockProvider';
/**
 * Sets global provider to a mock Provider
 *
 * @export
 * @class MgtMockProvider
 * @extends {LitElement}
 */
@registeredComponent('mgt-mock-provider')
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
