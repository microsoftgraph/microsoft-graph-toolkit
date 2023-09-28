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
        if (this._globalProvider.isMultiAccountSupportedAndEnabled) {
          this._globalProvider.removeActiveAccountChangedHandler(this.handleActiveAccountChanged);
        }
      }

      if (provider) {
        provider.onStateChanged(this.handleProviderStateChanged);
        if (provider.isMultiAccountSupportedAndEnabled) {
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
  public static onActiveAccountChanged(event: EventHandler<unknown>) {
    this._activeAccountChangedDispatcher.add(event);
  }

  /**
   * Remove event handler
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
   */
  public static removeActiveAccountChangedListener(event: EventHandler<unknown>) {
    this._activeAccountChangedDispatcher.remove(event);
  }

  /**
   * Gets the current signed in user
   *
   * @static
   * @memberof Providers
   */
  public static me(): Promise<User> {
    if (!this.client) {
      this._mePromise = null;
      return null;
    }

    // eslint-disable-next-line @typescript-eslint/no-misused-promises
    if (!this._mePromise) {
      this._mePromise = this.getMe();
    }

    return this._mePromise;
  }

  /**
   * Get current signed in user details
   *
   * @private
   * @static
   * @return {*}  {Promise<User>}
   * @memberof Providers
   */
  private static async getMe(): Promise<User> {
    try {
      const response: User = (await this.client.api('me').get()) as User;
      if (response?.id) {
        return response;
      }
    } catch {
      // no-op
    }

    return null;
  }

  /**
   * Gets the cache ID, creates one if it does not exist
   *
   * @static
   * @memberof Providers
   */
  public static async getCacheId() {
    if (this._cacheId) {
      return this._cacheId;
    }
    if (Providers.globalProvider?.state === ProviderState.SignedIn) {
      if (!this._cacheId) {
        const client = this.client;
        if (client) {
          try {
            this._cacheId = await this.createCacheId();
          } catch {
            // no-op
          }
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
  private static unsetCacheId() {
    this._cacheId = null;
    this._mePromise = null;
  }

  /**
   * Create cache ID
   *
   * @private
   * @static
   * @return {*}  {Promise<string>}
   * @memberof Providers
   */
  private static async createCacheId(): Promise<string> {
    if (Providers.globalProvider.isMultiAccountSupportedAndEnabled) {
      const cacheId = this.createCacheIdWithAccountDetails();
      if (cacheId) {
        return cacheId;
      }
    }
    return await this.createCacheIdWithUserDetails();
  }

  /**
   * Create a cache ID with user userID and principal name
   *
   * @static
   * @param {User} response
   * @return {*}
   * @memberof Providers
   */
  private static async createCacheIdWithUserDetails(): Promise<string> {
    const response: User = await this.me();
    if (response?.id) {
      return response.id + '-' + response.userPrincipalName;
    } else return null;
  }

  /**
   * Create cache ID with tenant ID and user ID
   *
   * @private
   * @static
   * @return {*}  {string}
   * @memberof Providers
   */
  private static createCacheIdWithAccountDetails(): string {
    const user = Providers.globalProvider.getActiveAccount();
    if (user.tenantId && user.id) {
      return user.tenantId + user.id;
    } else return null;
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

  private static readonly _eventDispatcher = new EventDispatcher<ProvidersChangedState>();

  private static readonly _activeAccountChangedDispatcher = new EventDispatcher<unknown>();

  private static _globalProvider: IProvider;
  private static _cacheId: string;
  private static _mePromise: Promise<User>;

  private static readonly handleProviderStateChanged = () => {
    if (!Providers.globalProvider || Providers.globalProvider.state !== ProviderState.SignedIn) {
      // clear current signed in user info
      Providers._mePromise = null;
    }

    Providers._eventDispatcher.fire(ProvidersChangedState.ProviderStateChanged);
  };

  private static readonly handleActiveAccountChanged = () => {
    Providers.unsetCacheId();
    Providers._activeAccountChangedDispatcher.fire(null);
  };
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
