/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement } from 'lit-element';
import { MockProvider } from './MockProvider';
import { Providers } from '../Providers';

@customElement('mgt-mock-provider')
export class MgtMockProvider extends LitElement {
  constructor() {
    super();
    Providers.globalProvider = new MockProvider(true);
  }
}
