import { CacheItem, CacheSchema, CacheService, CacheStore, schemas } from '@microsoft/mgt-react';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { cacheEntryIsValid } from './cacheEntryIsValid';

type CachedSubscriptionData = CacheItem & {
  chatId: string;
  subscriptions: Subscription[];
};

const buildCacheKey = (chatId: string, sessionId: string): string => `${chatId}:${sessionId}`;

export class SubscriptionsCache {
  private get cache(): CacheStore<CachedSubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedSubscriptionData>(conversation, conversation.stores.subscriptions);
  }

  public async loadSubscriptions(chatId: string, sessionId: string): Promise<CachedSubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      const data = await this.cache.getValue(buildCacheKey(chatId, sessionId));
      if (cacheEntryIsValid(data)) return data ?? undefined;
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
        subscriptions: [subscriptionRecord]
      };
    }

    await this.cache.putValue(buildCacheKey(chatId, sessionId), cacheEntry);
  }

  public async deleteCachedSubscriptions(chatId: string, sessionId: string): Promise<void> {
    await this.cache.delete(buildCacheKey(chatId, sessionId));
  }
}
