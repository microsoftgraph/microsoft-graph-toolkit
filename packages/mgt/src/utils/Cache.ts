// TODO ands consideration
// Should db clear on sign out
// SHould db be named with the user id

import { openDB } from 'idb';

export interface CacheSchema {
  version: number;
  name: string;
  stores: { [name: string]: CacheSchemaStore };
}

export interface CacheSchemaStore {
  key?: string;
}

export interface CacheItem {
  timeCached?: number;
}

export class Cache<T extends CacheItem> {
  private schema: CacheSchema;
  private store: string;

  public constructor(schema: CacheSchema, store: string) {
    if (!(store in schema.stores)) {
      throw Error('"store" must be defined in the "schema"');
    }

    this.schema = schema;
    this.store = store;
  }

  public async getValue(key: string): Promise<T> {
    if (!window.indexedDB) {
      console.log("browser doesn't support indexedDB");
      return null;
    }

    return (await this.getDb()).get(this.store, key);
  }

  public async putValue(key: string, item: T) {
    if (!window.indexedDB) {
      console.log("browser doesn't support indexedDB");
      return;
    }

    await (await this.getDb()).put(this.store, { ...item, timeCached: Date.now() }, key);
  }

  private getDb() {
    return openDB(this.getDBName(), this.schema.version, {
      upgrade: (db, oldVersion, newVersion, transaction) => {
        for (const storeName in this.schema.stores) {
          if (this.schema.stores.hasOwnProperty(storeName)) {
            db.createObjectStore(storeName);
            // TODO: setup key, index, etc
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
