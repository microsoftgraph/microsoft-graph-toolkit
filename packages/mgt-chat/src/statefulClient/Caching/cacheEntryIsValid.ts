/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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

export const cacheEntryIsValid = (cacheEntry: CacheItem | null) =>
  cacheEntry?.timeCached && getConversationCacheInvalidationTime() > Date.now() - cacheEntry.timeCached;
