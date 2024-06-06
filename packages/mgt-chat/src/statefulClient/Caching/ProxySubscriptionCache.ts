/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheSchema, CacheService, CacheStore, log, schemas } from '@microsoft/mgt-react';
import { ProxySubscription } from '../MGTProxyOperations';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { IDBPObjectStore } from 'idb';
import { ComponentType } from './SubscriptionCache';

type CachedProxySubscriptionData = CacheItem & {
  chatId: string;
  proxySubscriptions: ProxySubscription[];
  lastAccessDateTime: string;
  componentType: ComponentType;
};

export class ProxySubscriptionCache {
  private get cache(): CacheStore<CachedProxySubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedProxySubscriptionData>(conversation, conversation.stores.subscriptions);
  }

  public async loadProxySubscriptions(chatId: string): Promise<CachedProxySubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      let data;
      await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
        data = (await store.get(chatId)) as CachedProxySubscriptionData | undefined;
        if (data) {
          data.lastAccessDateTime = new Date().toISOString();
          await store.put(data, chatId);
        }
      });
      return data || undefined;
    }
    return undefined;
  }

  public async cacheProxySubscription(
    chatId: string,
    componentType: ComponentType,
    proxySubscriptionRecord: ProxySubscription
  ): Promise<void> {
    await this.cache.transaction(async (store: IDBPObjectStore<unknown, [string], string, 'readwrite'>) => {
      log('cacheSubscription', proxySubscriptionRecord);

      let cacheEntry = (await store.get(chatId)) as CachedProxySubscriptionData | undefined;
      if (cacheEntry && cacheEntry.chatId === chatId) {
        const subIndex = cacheEntry.proxySubscriptions.findIndex(
          s => s.subscription?.resource === proxySubscriptionRecord.subscription?.resource
        );
        if (subIndex !== -1) {
          cacheEntry.proxySubscriptions[subIndex] = proxySubscriptionRecord;
        } else {
          cacheEntry.proxySubscriptions.push(proxySubscriptionRecord);
        }
      } else {
        cacheEntry = {
          chatId,
          proxySubscriptions: [proxySubscriptionRecord],
          // we're cheating a bit here to ensure that we have a defined lastAccessDateTime
          // but we're updating the value for all cases before storing it.
          lastAccessDateTime: '',
          componentType
        };
      }
      cacheEntry.lastAccessDateTime = new Date().toISOString();

      await store.put(cacheEntry, chatId);
    });
  }

  public deleteCachedProxySubscriptions(chatId: string): Promise<void> {
    return this.cache.delete(chatId);
  }

  public loadInactiveProxySubscriptions(
    inactivityThreshold: string,
    componentType: ComponentType
  ): Promise<CachedProxySubscriptionData[]> {
    return this.cache
      .queryDb('lastAccessDateTime', IDBKeyRange.upperBound(inactivityThreshold))
      .then((data: CachedProxySubscriptionData[]) => data.filter(d => d.componentType === componentType))
      .catch(err => {
        // propogate the error back to the call of this function
        throw err;
      });
  }
}
