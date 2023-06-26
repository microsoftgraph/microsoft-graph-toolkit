/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Providers } from '../providers/Providers';
import { ProviderState } from '../providers/IProvider';
import { CacheStore } from './CacheStore';

/**
 * Localstorage key for storing names of cache databases
 *
 * @type {string}
 *
 */
export const dbListKey = 'mgt-db-list';

/**
 * Holds the cache options for cache store
 *
 * @export
 * @interface CacheConfig
 */
export interface CacheConfig {
  /**
   * Default global invalidation period
   *
   * @type {number}
   * @memberof CacheConfig
   */
  defaultInvalidationPeriod: number;
  /**
   * Controls whether the cache is enabled globally
   *
   * @type {boolean}
   * @memberof CacheConfig
   */
  isEnabled: boolean;
  /**
   * Cache options for groups store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  groups: CacheOptions;
  /**
   * Cache options for people store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  people: CacheOptions;
  /**
   * Cache options for photos store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  photos: CacheOptions;
  /**
   * Cache options for presence store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */

  presence: CacheOptions;
  /**
   * Cache options for users store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  users: CacheOptions;

  /**
   * Cache options for a generic response store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  response: CacheOptions;

  /**
   * Cache options for files store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  files: CacheOptions;
  /**
   * Cache options for fileLists store
   *
   * @type {CacheOptions}
   * @memberof CacheConfig
   */
  fileLists: CacheOptions;
}

/**
 * Options for each store
 *
 * @export
 * @interface CacheOptions
 */
export interface CacheOptions {
  /**
   * Defines the time (in ms) for objects in the store to expire
   *
   * @type {number}
   * @memberof CacheOptions
   */
  invalidationPeriod: number;
  /**
   * Whether the store is enabled or not
   *
   * @type {boolean}
   * @memberof CacheOptions
   */
  isEnabled: boolean;
}

/**
 * class in charge of managing all the caches and their stores
 *
 * @export
 * @class CacheService
 */
export class CacheService {
  /**
   * Looks for existing cache, otherwise creates a new one
   *
   * @static
   * @template T
   * @param {CacheSchema} schema
   * @param {string} storeName
   * @returns {CacheStore<T>}
   * @memberof CacheService
   */
  public static getCache<T extends CacheItem>(schema: CacheSchema, storeName: string): CacheStore<T> {
    const key = `${schema.name}/${storeName}`;

    if (!this.isInitialized) {
      this.init();
    }

    if (!this.cacheStore.has(storeName)) {
      this.cacheStore.set(key, new CacheStore<T>(schema, storeName));
    }
    return this.cacheStore.get(key) as CacheStore<T>;
  }

  /**
   * Clears cache for a single user when ID is passed
   *
   * @static
   * @param {string} id
   * @memberof CacheService
   */
  public static clearCacheById(id: string): Promise<unknown> {
    const work: Promise<void>[] = [];
    const oldDbArray: string[] = JSON.parse(localStorage.getItem(dbListKey)) as string[];
    if (oldDbArray) {
      const newDbArray: string[] = [];
      oldDbArray.forEach(x => {
        if (x.includes(id)) {
          work.push(
            new Promise<void>((resolve, reject) => {
              const delReq = indexedDB.deleteDatabase(x);
              delReq.onsuccess = () => resolve();
              delReq.onerror = () => {
                console.error(`🦒: ${delReq.error.name} occurred deleting cache: ${x}`, delReq.error.message);
                reject();
              };
            })
          );
        } else {
          newDbArray.push(x);
        }
      });
      if (newDbArray.length > 0) {
        localStorage.setItem(dbListKey, JSON.stringify(newDbArray));
      } else {
        localStorage.removeItem(dbListKey);
      }
    }
    return Promise.all(work);
  }

  private static readonly cacheStore = new Map<string, CacheStore<CacheItem>>();
  private static isInitialized = false;

  private static readonly cacheConfig: CacheConfig = {
    defaultInvalidationPeriod: 3600000,
    groups: {
      invalidationPeriod: null,
      isEnabled: true
    },
    isEnabled: true,
    people: {
      invalidationPeriod: null,
      isEnabled: true
    },
    photos: {
      invalidationPeriod: null,
      isEnabled: true
    },
    presence: {
      invalidationPeriod: 300000,
      isEnabled: true
    },
    users: {
      invalidationPeriod: null,
      isEnabled: true
    },
    response: {
      invalidationPeriod: null,
      isEnabled: true
    },
    files: {
      invalidationPeriod: null,
      isEnabled: true
    },
    fileLists: {
      invalidationPeriod: null,
      isEnabled: true
    }
  };

  /**
   * returns the cacheconfig object
   *
   * @readonly
   * @static
   * @type {CacheConfig}
   * @memberof CacheService
   */
  public static get config(): CacheConfig {
    return this.cacheConfig;
  }

  /**
   * Checks for current sign in state and see if it has changed from signed-in to signed out
   *
   *
   * @private
   * @static
   * @memberof CacheService
   */
  private static init() {
    let previousState: ProviderState;
    if (Providers.globalProvider) {
      previousState = Providers.globalProvider.state;
    }

    // eslint-disable-next-line @typescript-eslint/no-misused-promises
    Providers.onProviderUpdated(async () => {
      if (previousState === ProviderState.SignedIn && Providers.globalProvider.state === ProviderState.SignedOut) {
        const id = await Providers.getCacheId();
        if (id !== null) {
          await this.clearCacheById(id);
        }
      }
      previousState = Providers.globalProvider.state;
    });
    this.isInitialized = true;
  }
}

/**
 * Represents organization for a cache
 *
 * @export
 * @interface CacheSchema
 */
export interface CacheSchema {
  /**
   * version number of cache, useful for upgrading
   *
   * @type {number}
   * @memberof CacheSchema
   */
  version: number;
  /**
   * name of the cache
   *
   * @type {string}
   * @memberof CacheSchema
   */
  name: string;
  /**
   * list of stores in the cache
   *
   * @type {{ [name: string]: CacheSchemaStore }}
   * @memberof CacheSchema
   */
  stores: Record<string, string>;
}

/**
 * item that is stored in cache
 *
 * @export
 * @interface CacheItem
 */
export interface CacheItem {
  /**
   * date and time that item was retrieved from api/stored in cache
   *
   * @type {number}
   * @memberof CacheItem
   */
  timeCached?: number;
}
