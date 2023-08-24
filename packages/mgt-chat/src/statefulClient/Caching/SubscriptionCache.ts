import { CacheItem, CacheSchema, CacheService, CacheStore, schemas } from '@microsoft/mgt-react';
import { Subscription } from '@microsoft/microsoft-graph-types';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { cacheEntryIsValid } from './cacheEntryIsValid';

type CachedSubscriptionData = CacheItem & {
  chatId: string;
  subscriptions: Subscription[];
};

const subscriptionCacheKey = 'graph-current-subscriptions';

export class SubscriptionsCache {
  private get cache(): CacheStore<CachedSubscriptionData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedSubscriptionData>(conversation, conversation.stores.chats);
  }

  public async loadSubscriptions(): Promise<CachedSubscriptionData | undefined> {
    if (isConversationCacheEnabled()) {
      const data = await this.cache.getValue(subscriptionCacheKey);
      if (cacheEntryIsValid(data)) return data ?? undefined;
    }
    return undefined;
  }

  public async cacheSubscription(chatId: string, subscriptionRecord: Subscription): Promise<void> {
    let cacheEntry = await this.loadSubscriptions();
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

    await this.cache.putValue(subscriptionCacheKey, cacheEntry);
  }

  public async clearCachedSubscriptions(): Promise<void> {
    await this.cache.delete(subscriptionCacheKey);
  }
}
