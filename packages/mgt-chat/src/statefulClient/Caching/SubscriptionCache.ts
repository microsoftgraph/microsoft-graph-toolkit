/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheSchema, CacheService, CacheStore, schemas } from '@microsoft/mgt-react';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { cacheEntryIsValid } from './cacheEntryIsValid';

type CachedSubscriptionData = CacheItem & {
  chatId: string;
  sessionId: string;
  subscriptions: Subscription[];
  lastAccessDateTime: string;
};

const buildCacheKey = (chatId: string, sessionId: string): string => `${chatId}:${sessionId}`;

export class SubscriptionsCache {
  private get cache(): CacheStore<CachedSubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedSubscriptionData>(conversation, conversation.stores.subscriptions);
  }

  public async loadSubscriptions(chatId: string, sessionId: string): Promise<CachedSubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      const cacheKey = buildCacheKey(chatId, sessionId);
      const data = await this.cache.getValue(cacheKey);
      if (data && cacheEntryIsValid(data)) {
        data.lastAccessDateTime = new Date().toISOString();
        await this.cache.putValue(cacheKey, data);
        return data;
      }
    }
    return undefined;
  }

  public async cacheSubscription(chatId: string, sessionId: string, subscriptionRecord: Subscription): Promise<void> {
    let cacheEntry = await this.loadSubscriptions(chatId, sessionId);
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
        sessionId,
        subscriptions: [subscriptionRecord],
        // we're cheating a bit here to ensure that we have a defined lastAccessDateTime
        // but we're updating the value for all cases before storing it.
        lastAccessDateTime: ''
      };
    }
    cacheEntry.lastAccessDateTime = new Date().toISOString();

    await this.cache.putValue(buildCacheKey(chatId, sessionId), cacheEntry);
  }

  public deleteCachedSubscriptions(chatId: string, sessionId: string): Promise<void> {
    return this.cache.delete(buildCacheKey(chatId, sessionId));
  }

  public loadInactiveSubscriptions(inactivityThreshold: string): Promise<CachedSubscriptionData[]> {
    return this.cache.queryDb('lastAccessDateTime', IDBKeyRange.upperBound(inactivityThreshold));
  }
}
