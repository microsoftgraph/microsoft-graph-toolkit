/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Providers } from '../../Providers';
import { MgtBaseComponent } from '../baseComponent';

export abstract class MgtBaseProvider extends MgtBaseComponent {
  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadState());
    this.loadState();
  }

  private async loadState() {
    const provider = Providers.globalProvider;

    if (provider) {
      provider.onStateChanged(() => {
        this.fireCustomEvent('onStateChanged', provider.state);
      });
    }
  }
}
