/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher, EventHandler, IProvider } from './providers/IProvider';

export class Providers {
  public static get globalProvider(): IProvider {
    return this._globalProvider;
  }

  public static set globalProvider(provider: IProvider) {
    if (provider !== this._globalProvider) {
      if (this._globalProvider) {
        this._globalProvider.removeStateChangedHandler(this.handleProviderStateChanged);
      }

      if (provider) {
        provider.onStateChanged(this.handleProviderStateChanged);
      }

      this._globalProvider = provider;
      this._eventDispatcher.fire(ProvidersChangedState.ProviderChanged);
    }
  }

  public static onProviderUpdated(event: EventHandler<ProvidersChangedState>) {
    this._eventDispatcher.add(event);
  }

  public static removeProviderUpdatedListener(event: EventHandler<ProvidersChangedState>) {
    this._eventDispatcher.remove(event);
  }
  private static _eventDispatcher: EventDispatcher<ProvidersChangedState> = new EventDispatcher<
    ProvidersChangedState
  >();
  private static _globalProvider: IProvider;

  private static handleProviderStateChanged() {
    Providers._eventDispatcher.fire(ProvidersChangedState.ProviderStateChanged);
  }
}

export enum ProvidersChangedState {
  ProviderChanged,
  ProviderStateChanged
}
