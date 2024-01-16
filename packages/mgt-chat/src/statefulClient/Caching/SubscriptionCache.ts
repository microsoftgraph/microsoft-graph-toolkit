/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheSchema, CacheService, CacheStore, log, schemas } from '@microsoft/mgt-react';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { IDBPObjectStore } from 'idb';

export enum ComponentType {
  User = 'user',
  Chat = 'chat'
}

type CachedSubscriptionData = CacheItem & {
  componentEntityId: string;
  sessionId: string;
  subscriptions: Subscription[];
  lastAccessDateTime: string;
  componentType: ComponentType;
};

const buildCacheKey = (componentEntityId: string, sessionId: string): string => `${componentEntityId}:${sessionId}`;

export class SubscriptionsCache {
  private get cache(): CacheStore<CachedSubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedSubscriptionData>(conversation, conversation.stores.subscriptions);
  }

  public async loadSubscriptions(
    componentEntityId: string,
    sessionId: string
  ): Promise<CachedSubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      const cacheKey = buildCacheKey(componentEntityId, sessionId);
      let data;
      await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
        data = (await store.get(cacheKey)) as CachedSubscriptionData | undefined;
        if (data) {
          data.lastAccessDateTime = new Date().toISOString();
          await store.put(data, cacheKey);
        }
      });
      return data || undefined;
    }
    return undefined;
  }

  public async cacheSubscription(
    componentEntityId: string,
    componentType: ComponentType,
    sessionId: string,
    subscriptionRecord: Subscription
  ): Promise<void> {
    await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
      log('cacheSubscription', subscriptionRecord);
      const cacheKey = buildCacheKey(componentEntityId, sessionId);

      let cacheEntry = (await store.get(cacheKey)) as CachedSubscriptionData | undefined;
      if (cacheEntry && cacheEntry.componentEntityId === componentEntityId) {
        const subIndex = cacheEntry.subscriptions.findIndex(s => s.resource === subscriptionRecord.resource);
        if (subIndex !== -1) {
          cacheEntry.subscriptions[subIndex] = subscriptionRecord;
        } else {
          cacheEntry.subscriptions.push(subscriptionRecord);
        }
      } else {
        cacheEntry = {
          componentEntityId,
          componentType,
          sessionId,
          subscriptions: [subscriptionRecord],
          // we're cheating a bit here to ensure that we have a defined lastAccessDateTime
          // but we're updating the value for all cases before storing it.
          lastAccessDateTime: ''
        };
      }
      cacheEntry.lastAccessDateTime = new Date().toISOString();

      await store.put(cacheEntry, cacheKey);
    });
  }

  public deleteCachedSubscriptions(componentEntityId: string, sessionId: string): Promise<void> {
    return this.cache.delete(buildCacheKey(componentEntityId, sessionId));
  }

  public loadInactiveSubscriptions(
    inactivityThreshold: string,
    componentType: ComponentType
  ): Promise<CachedSubscriptionData[]> {
    return this.cache
      .queryDb('lastAccessDateTime', IDBKeyRange.upperBound(inactivityThreshold))
      .then((data: CachedSubscriptionData[]) => data.filter(d => d.componentType === componentType))
      .catch(err => {
        // propogate the error back to the call of this function
        throw err;
      });
  }
}
