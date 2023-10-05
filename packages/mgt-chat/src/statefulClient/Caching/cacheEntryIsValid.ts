import { CacheItem, CacheOptions, CacheService } from '@microsoft/mgt-react';

/**
 * Defines the expiration time
 */
const getConversationCacheInvalidationTime = (): number => {
  const conversation = CacheService.config.conversation as CacheOptions;
  return conversation.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;
};

export const cacheEntryIsValid = (cacheEntry: CacheItem | null) =>
  cacheEntry?.timeCached && getConversationCacheInvalidationTime() > Date.now() - cacheEntry.timeCached;
