/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { User } from '@microsoft/microsoft-graph-types';

import { EventDispatcher, EventHandler } from '../utils/EventDispatcher';
import { IProvider, ProviderState } from './IProvider';

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
        if (this._globalProvider.isMultiAccountEnabled) {
          this._globalProvider.removeActiveAccountChangedHandler(this.handleActiveAccountChanged);
        }
      }

      if (provider) {
        provider.onStateChanged(this.handleProviderStateChanged);
        if (provider.isMultiAccountEnabled) {
          provider.onActiveAccountChanged(this.handleActiveAccountChanged);
        }
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

  /**
   * Fires event when Provider changes state
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
   */
  public static onActiveAccountChanged(event: EventHandler<any>) {
    this._activeAccountChangedDispatcher.add(event);
  }

  /**
   * Remove event handler
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
   */
  public static removeActiveAccountChangedListener(event: EventHandler<any>) {
    this._activeAccountChangedDispatcher.remove(event);
  }

  /**
   * Gets the current signed in user
   *
   * @static
   * @memberof Providers
   */
  public static async me() {
    if (!this._me) {
      const client = this.client;
      if (client) {
        try {
          const response: User = await client.api('me').get();
          if (response && response.id) {
            this._me = response;
          }
        } catch {}
      }
    }

    return this._me;
  }

  /**
   * Gets the cache ID, creates one if it does not exist
   *
   * @static
   * @memberof Providers
   */
  public static async getCacheId(force?: boolean) {
    if (this._cacheId && !force) {
      return this._cacheId;
    }
    if (Providers.globalProvider && Providers.globalProvider.state == ProviderState.SignedIn) {
      if (!this._cacheId || force) {
        this._cacheId = null;
        const client = this.client;
        if (client) {
          try {
            const response: User = await client.api('me').get();
            if (response && response.id) {
              this._cacheId = this.createCacheId(response);
            }
          } catch {}
        }
      }
    }
    return this._cacheId;
  }

  /**
   * Unset the cache ID
   *
   * @static
   * @memberof Providers
   */
  public static unsetCacheId() {
    this._cacheId = null;
  }

  /**
   * Create a cache ID
   *
   * @static
   * @param {User} response
   * @return {*}
   * @memberof Providers
   */
  public static createCacheId(response: User) {
    return response.id + '-' + response.userPrincipalName;
  }

  /**
   * Gets the current graph client
   *
   * @readonly
   * @static
   * @type {Client}
   * @memberof Providers
   */
  public static get client(): Client {
    if (Providers.globalProvider && Providers.globalProvider.state === ProviderState.SignedIn) {
      return Providers.globalProvider.graph.client;
    }
    return null;
  }

  private static _eventDispatcher: EventDispatcher<ProvidersChangedState> = new EventDispatcher<ProvidersChangedState>();

  private static _activeAccountChangedDispatcher: EventDispatcher<any> = new EventDispatcher<any>();

  private static _globalProvider: IProvider;
  private static _me: User;
  private static _cacheId: string;

  private static handleProviderStateChanged() {
    if (!Providers.globalProvider || Providers.globalProvider.state !== ProviderState.SignedIn) {
      // clear current signed in user info
      Providers._me = null;
    }

    Providers._eventDispatcher.fire(ProvidersChangedState.ProviderStateChanged);
  }

  private static handleActiveAccountChanged() {
    Providers._activeAccountChangedDispatcher.fire(null);
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
