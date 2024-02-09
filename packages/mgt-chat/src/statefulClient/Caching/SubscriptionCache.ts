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

type CachedSubscriptionData = CacheItem & {
  chatId: string;
  subscriptions: Subscription[];
  lastAccessDateTime: string;
};

export class SubscriptionsCache {
  private get cache(): CacheStore<CachedSubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedSubscriptionData>(conversation, conversation.stores.subscriptions);
  }

  public async loadSubscriptions(chatId: string): Promise<CachedSubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      let data;
      await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
        data = (await store.get(chatId)) as CachedSubscriptionData | undefined;
        if (data) {
          data.lastAccessDateTime = new Date().toISOString();
          await store.put(data, chatId);
        }
      });
      return data || undefined;
    }
    return undefined;
  }

  public async cacheSubscription(chatId: string, subscriptionRecord: Subscription): Promise<void> {
    await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
      log('cacheSubscription', subscriptionRecord);

      let cacheEntry = (await store.get(chatId)) as CachedSubscriptionData | undefined;
      if (cacheEntry && cacheEntry.chatId === chatId) {
        const subIndex = cacheEntry.subscriptions.findIndex(s => s.resource === subscriptionRecord.resource);
        if (subIndex !== -1) {
          cacheEntry.subscriptions[subIndex] = subscriptionRecord;
        } else {
          cacheEntry.subscriptions.push(subscriptionRecord);
        }
      } else {
        cacheEntry = {
          chatId,
          subscriptions: [subscriptionRecord],
          // we're cheating a bit here to ensure that we have a defined lastAccessDateTime
          // but we're updating the value for all cases before storing it.
          lastAccessDateTime: ''
        };
      }
      cacheEntry.lastAccessDateTime = new Date().toISOString();

      await store.put(cacheEntry, chatId);
    });
  }

  public deleteCachedSubscriptions(chatId: string): Promise<void> {
    return this.cache.delete(chatId);
  }

  public loadInactiveSubscriptions(inactivityThreshold: string): Promise<CachedSubscriptionData[]> {
    return this.cache.queryDb('lastAccessDateTime', IDBKeyRange.upperBound(inactivityThreshold));
  }
}
