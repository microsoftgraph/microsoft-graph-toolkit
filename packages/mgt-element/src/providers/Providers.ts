/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IProvider } from './IProvider';
import { EventDispatcher, EventHandler } from '../utils/EventDispatcher';

/**
 * Provides implementation for acquiring the necessary access token for calling the Microsoft Graph APIs.
 *
 * @export
 * @class Providers
 */
export class Providers {
  /**
   * returns the value of provider used globally. All components use this property to get a reference to the provider.
   *
   * @static
   * @type {IProvider}
   * @memberof Providers
   */
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

  /**
   * Fires event when Provider changes state
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
   */
  public static onProviderUpdated(event: EventHandler<ProvidersChangedState>) {
    this._eventDispatcher.add(event);
  }
  /**
   * Remove event handler
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
   */
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

/**
 * on Provider Change State
 *
 * @export
 * @enum {number}
 */
export enum ProvidersChangedState {
  /**
   * ProviderChanged = 0
   */
  ProviderChanged,
  /**
   * ProviderStateChanged = 1
   */
  ProviderStateChanged
}
