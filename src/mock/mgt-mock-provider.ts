/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, LitElement } from 'lit-element';
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
  constructor() {
    super();
    Providers.globalProvider = new MockProvider(true);
  }
}
