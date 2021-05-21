/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { openDB } from 'idb';
import { Providers } from '../providers/Providers';
import { ProviderState } from '../providers/IProvider';

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
// tslint:disable-next-line: max-classes-per-file
export class CacheService {
  /**
   *  Looks for existing cache, otherwise creates a new one
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
   * Clears all the stores within the cache
   */
  public static clearCaches() {
    this.cacheStore.forEach(x => indexedDB.deleteDatabase(x.getDBName()));
  }

  private static cacheStore: Map<string, CacheStore<CacheItem>> = new Map();
  private static isInitialized: boolean = false;

  private static cacheConfig: CacheConfig = {
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

    Providers.onProviderUpdated(() => {
      if (previousState === ProviderState.SignedIn && Providers.globalProvider.state === ProviderState.SignedOut) {
        this.clearCaches();
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
  stores: { [name: string]: string };
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

/**
 * Represents a store in the cache
 *
 * @class CacheStore
 * @template T
 */
// tslint:disable-next-line: max-classes-per-file
export class CacheStore<T extends CacheItem> {
  private schema: CacheSchema;
  private store: string;

  public constructor(schema: CacheSchema, store: string) {
    if (!(store in schema.stores)) {
      throw Error('"store" must be defined in the "schema"');
    }

    this.schema = schema;
    this.store = store;
  }

  /**
   * gets value from cache for the given key
   *
   * @param {string} key
   * @returns {Promise<T>}
   * @memberof Cache
   */
  public async getValue(key: string): Promise<T> {
    if (!window.indexedDB) {
      return null;
    }
    try {
      return (await this.getDb()).get(this.store, key);
    } catch (e) {
      return null;
    }
  }

  /**
   * inserts value into cache for the given key
   *
   * @param {string} key
   * @param {T} item
   * @returns
   * @memberof Cache
   */
  public async putValue(key: string, item: T) {
    if (!window.indexedDB) {
      return;
    }
    try {
      await (await this.getDb()).put(this.store, { ...item, timeCached: Date.now() }, key);
    } catch (e) {
      return;
    }
  }

  /**
   * Clears the store of all stored values
   *
   * @returns
   * @memberof Cache
   */
  public async clearStore() {
    if (!window.indexedDB) {
      return;
    }
    try {
      (await this.getDb()).clear(this.store);
    } catch (e) {
      return;
    }
  }

  /**
   * Returns the name of the parent DB that the cache store belongs to
   */
  public getDBName() {
    return `mgt-${this.schema.name}`;
  }

  private getDb() {
    return openDB(this.getDBName(), this.schema.version, {
      upgrade: (db, oldVersion, newVersion, transaction) => {
        for (const storeName in this.schema.stores) {
          if (this.schema.stores.hasOwnProperty(storeName)) {
            db.objectStoreNames.contains(storeName) || db.createObjectStore(storeName);
          }
        }
      }
    });
  }
}
