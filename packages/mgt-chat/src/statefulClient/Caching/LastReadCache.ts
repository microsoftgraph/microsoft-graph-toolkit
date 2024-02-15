/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheSchema, CacheService, CacheStore, Providers, schemas } from '@microsoft/mgt-react';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { cacheEntryIsValid } from './cacheEntryIsValid';

/*
 * This implementation does not store the last read time in the "chats" store by design. Records
 * in the "chats" store potentially contain hundreds of messages and can be quite large. However,
 * those records are only read or written when the chat is loaded or when the user requests more
 * historical messages. The last read time is updated on a timer that by default is configured
 * for every 30 seconds, and this would put more pressure on the IndexedDB storage than is needed.
 */

interface LastReadData extends CacheItem {
  chatId: string;
  lastReadTime: string; // ISO 8601 timestamp
}

export class LastReadCache {
  private get cache(): CacheStore<LastReadData> {
    const conversation: CacheSchema = schemas.conversation;
    const cache = CacheService.getCache<LastReadData>(conversation, conversation.stores.lastRead);
    cache.getDBName = async () => {
      const id = await Providers.getCacheId();
      if (id) {
        // this removal of dashes is done because when a user signs out, the code looks for contains matches
        // from getCacheId to remove the cache/db. If the id has no dashes, it will not be removed.
        return `mgt-lastreadcache-${id.replace(/-/g, '')}`;
      }

      // needed for build error
      throw new Error('CacheId is not available');
    };
    return cache;
  }

  public async loadLastReadTime(chatId: string): Promise<LastReadData | null | undefined> {
    if (isConversationCacheEnabled()) {
      const data = await this.cache.getValue(chatId);
      if (data && cacheEntryIsValid(data)) return data;
    }
    return undefined;
  }

  public async cacheLastReadTime(chatId: string, lastReadTime: Date | string): Promise<LastReadData> {
    if (lastReadTime instanceof Date) {
      lastReadTime = lastReadTime.toISOString();
    }
    const data = { chatId, lastReadTime };
    await this.cache.putValue(chatId, data);
    return data;
  }
}
