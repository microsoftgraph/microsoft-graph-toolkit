import { CacheItem, CacheService } from '@microsoft/mgt-react';

/**
 * Defines the expiration time
 */
const getConversationCacheInvalidationTime = (): number => {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const conversation = CacheService.config.conversation;
  // eslint-disable-next-line @typescript-eslint/no-unsafe-return, @typescript-eslint/no-unsafe-member-access
  return conversation.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;
};

export const cacheEntryIsValid = (cacheEntry: CacheItem) =>
  cacheEntry.timeCached && getConversationCacheInvalidationTime() > Date.now() - cacheEntry.timeCached;
