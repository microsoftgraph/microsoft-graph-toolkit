/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement, property } from 'lit-element';
import { Providers } from '../Providers';
import { MockProvider } from './MockProvider';
/**
 * Sets global provider to a mock Provider
 *
 * @export
 * @class MgtMockProvider
 * @extends {LitElement}
 */
@customElement('mgt-mock-provider')
export class MgtMockProvider extends LitElement {
  /**
   * A property to allow the developer to start the sample logged out if they desired.
   *
   * @memberof MgtMockProvider
   */
  @property({
    attribute: 'signed-in',
    type: Boolean
  })
  public signedIn;

  constructor() {
    super();

    // Access the 'signed-in' attribute directly.
    // LitElement doesn't parse attributes early enough for us enact on them from the constructor.
    const signedInVal = (<any>this).getAttribute('signed-in');
    this.signedIn = signedInVal !== 'false' && signedInVal !== null;

    Providers.globalProvider = new MockProvider(this.signedIn);
  }
}
