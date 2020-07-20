// TODO ands consideration
// Should db clear on sign out
// SHould db be named with the user id

import { openDB } from 'idb';
import { Providers } from '../Providers';
import { ProviderState } from '../providers/IProvider';

/**
 * class in charge of managing all the caches and their stores
 *
 * @export
 * @class CacheService
 */
export class CacheService {
  /**
   *  Looks for existing cache, otherwise creates a new one
   *
   * @static
   * @template T
   * @param {CacheSchema} schema
   * @param {string} store
   * @returns {Cache<T>}
   * @memberof CacheService
   */
  public static getCache<T extends CacheItem>(schema: CacheSchema, store: string): Cache<T> {
    const key = `${schema.name}/${store}`;

    if (!this.isInitialized) {
      this.init();
    }

    if (!this.cacheStore.has(store)) {
      this.cacheStore.set(key, new Cache<T>(schema, store));
    }
    return this.cacheStore.get(key) as Cache<T>;
  }
  private static cacheStore: Map<string, Cache<CacheItem>> = new Map();
  private static isInitialized: boolean = false;
  private static _state: ProviderState;

  /**
   * @private
   * @static
   * @memberof CacheService
   */
  private static init() {
    if (Providers.globalProvider) {
      this._state = Providers.globalProvider.state;
    }

    Providers.onProviderUpdated(() => {
      if (this._state === ProviderState.SignedIn && Providers.globalProvider.state === ProviderState.SignedOut) {
        this.clearCaches();
      }
      this._state = Providers.globalProvider.state;
    });
    this.isInitialized = true;
  }

  private static clearCaches() {
    this.cacheStore.forEach(x => x.clearStore());
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
  stores: { [name: string]: CacheSchemaStore };
}

/**
 * Represents an individual store within each cache
 *
 * @export
 * @interface CacheSchemaStore
 */
export interface CacheSchemaStore {
  /**
   * key used to access values in the cache
   *
   * @type {string}
   * @memberof CacheSchemaStore
   */
  key?: string;
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
 * Poorly named, but represents a store in the cache
 *
 * @class Cache
 * @template T
 */
// tslint:disable-next-line: max-classes-per-file
class Cache<T extends CacheItem> {
  private schema: CacheSchema;
  private store: string;
  private _state: ProviderState;

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
      console.log("browser doesn't support indexedDB");
      return null;
    }

    return (await this.getDb()).get(this.store, key);
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
      console.log("browser doesn't support indexedDB");
      return;
    }

    await (await this.getDb()).put(this.store, { ...item, timeCached: Date.now() }, key);
  }

  /**
   * clears the store of all stored values
   *
   * @returns
   * @memberof Cache
   */
  public async clearStore() {
    if (!window.indexedDB) {
      console.log("browser doesn't support indexedDB");
      return;
    }

    (await this.getDb()).clear(this.store);
  }

  private getDb() {
    return openDB(this.getDBName(), this.schema.version, {
      upgrade: (db, oldVersion, newVersion, transaction) => {
        for (const storeName in this.schema.stores) {
          if (this.schema.stores.hasOwnProperty(storeName)) {
            db.createObjectStore(storeName);
          }
        }
      }
    });
  }

  private getDBName() {
    // TODO: signed in user id
    return `mgt-${this.schema.name}`;
  }
}
