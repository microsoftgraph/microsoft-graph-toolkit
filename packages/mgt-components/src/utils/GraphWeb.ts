/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheService, IGraph, Providers } from "@microsoft/mgt-element";
import { User } from "@microsoft/microsoft-graph-types";
import { schemas } from "../graph/cacheStores";
import { validUserByIdScopes } from "../graph/graph.user";
interface CacheUserIDs extends CacheItem {
  hashedUsers: string;
}

const getUserInvalidationTime = (): number =>
  CacheService.config.users.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;


/**
 * Whether or not the cache is enabled
 */
export const getIsUsersCacheEnabled = (): boolean =>
  CacheService.config.users.isEnabled && CacheService.config.isEnabled;


export class GraphWeb {
  private readonly worker: SharedWorker;
  private readonly graph: IGraph = Providers.globalProvider?.graph.forComponent(this);

  constructor() {
    this.worker = new SharedWorker(
      /* webpackChunkName: "graph-web-worker" */ new URL('./graphWebWorker.js', import.meta.url)
    );
    console.log(this.graph)
    this.worker.port.onmessage = this.onMessage;
  }

  public close() {
    this.worker.port.close();
  }

  private readonly onMessage = (e: MessageEvent) => {
    console.log('message from worker', e.data);
  };

  public async getUsersForUserIds(userIds: string[], filters: string): Promise<void> {
    console.log('webWorker: getUsersForUserIds', userIds);
    console.log('filters', filters);
    // 1. Create a key from the userIds and filters, use key to check the cache
    const hashKey = await this.hashArray(userIds)
    console.log(hashKey)
    const cache = CacheService.getCache<CacheUserIDs>(schemas.users, schemas.users.stores.hashedUsers)
    const cacheRes = await cache.getValue(hashKey);
    if (cacheRes && getUserInvalidationTime() > Date.now() - cacheRes.timeCached) {
      console.log('cache hit', cacheRes)
      return
    }
    console.log('cache miss')
    // 2. If no users/ cache exists get the users from the graph and store them in the cache
    const batch = this.graph.createBatch<User>();
    for (const userId of userIds) {
      let apiUrl = `/users/${userId}`;
      if (filters) {
        apiUrl += `${apiUrl}?$filter=${filters}&$count=true`;
      }
      batch.get(userId, apiUrl, validUserByIdScopes, filters ? { ConsistencyLevel: 'eventual' } : {});
    }
    if (batch.hasRequests) {
      const users: User[] = [];
      const responses = await batch.executeAll();
      for (const id of userIds) {
        const response = responses.get(id);
        if (response?.content) {
          const user = response.content;
          users.push(user);
        }
      }
      console.table(users)
      await cache.putValue(hashKey, { hashedUsers: JSON.stringify(users) });
      // 3. Do nothing if the users are already cached
    };
  }

  private async hashArray(arr: string[]): Promise<string> {
    const encoder = new TextEncoder();
    const data = encoder.encode(JSON.stringify(arr));
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArrayBuffer = Array.from(new Uint8Array(hashBuffer));
    const hashHex = hashArrayBuffer.map(b => b.toString(16).padStart(2, '0')).join('');
    return hashHex;
  }
}
