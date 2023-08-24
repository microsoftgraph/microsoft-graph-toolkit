import { CacheItem, CacheService } from '@microsoft/mgt-react';

/**
 * Defines the expiration time
 */
const getConversationCacheInvalidationTime = (): number => {
  return CacheService.config.conversation.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;
};

export const cacheEntryIsValid = (cacheEntry: CacheItem | null) =>
  cacheEntry?.timeCached && getConversationCacheInvalidationTime() > Date.now() - cacheEntry.timeCached;
