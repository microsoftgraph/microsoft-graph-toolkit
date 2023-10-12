import { CacheItem, CacheSchema, CacheService, CacheStore, schemas } from '@microsoft/mgt-react';
import { ChatMessage } from '@microsoft/microsoft-graph-types';
import { GraphCollection } from '../graph.chat';
import { isConversationCacheEnabled } from './isConversationCacheEnabled';
import { cacheEntryIsValid } from './cacheEntryIsValid';

interface CachedMessageData extends CacheItem, GraphCollection<ChatMessage> {
  chatId: string;
  lastModifiedDateTime: string; // ISO 8601 timestamp, should be max from values in messages
}

export class MessageCache {
  private get cache(): CacheStore<CachedMessageData> {
    const conversation: CacheSchema = schemas.conversation;
    return CacheService.getCache<CachedMessageData>(conversation, conversation.stores.chats);
  }

  public async loadMessages(chatId: string): Promise<CachedMessageData | null | undefined> {
    if (isConversationCacheEnabled()) {
      const data = await this.cache.getValue(chatId);
      if (cacheEntryIsValid(data)) return data;
    }
    return undefined;
  }

  public async cacheMessages(
    chatId: string,
    messages: ChatMessage[],
    updateNextLink = false,
    nextLink?: string
  ): Promise<CachedMessageData> {
    const cachedData = await this.cache.getValue(chatId);
    if (!cachedData) {
      const maxLastModifiedDateTime = messages.reduce((acc: string, curr: ChatMessage) => {
        if (curr.lastModifiedDateTime && curr.lastModifiedDateTime > acc) acc = curr.lastModifiedDateTime;
        return acc;
      }, '');
      const data = {
        chatId,
        value: messages,
        lastModifiedDateTime: maxLastModifiedDateTime,
        nextLink
      };
      // nothing cached, create and store.
      await this.cache.putValue(chatId, data);
      return data;
    } else {
      // need to iterate through the messages and either push or splice them into the array.
      messages.forEach(m => this.addMessageToCacheData(m, cachedData));
      if (updateNextLink) {
        cachedData.nextLink = nextLink;
      }
      // update persisted data
      await this.cache.putValue(chatId, cachedData);
      return cachedData;
    }
  }

  public async cacheMessage(chatId: string, message: ChatMessage): Promise<void> {
    const cachedData = await this.cache.getValue(chatId);
    if (!cachedData) {
      // nothing cached, create and store.
      await this.cache.putValue(chatId, {
        chatId,
        value: [message],
        lastModifiedDateTime: message.lastModifiedDateTime || ''
      });
    } else {
      this.addMessageToCacheData(message, cachedData);
      // update persisted data
      await this.cache.putValue(chatId, cachedData);
    }
  }

  private addMessageToCacheData(message: ChatMessage, cachedData: CachedMessageData) {
    const spliceIndex = cachedData.value.findIndex(m => m.id === message.id);
    if (spliceIndex !== -1) {
      cachedData.value.splice(spliceIndex, 1, message);
    } else {
      cachedData.value.push(message);
    }
    // coerce potential nullish values to an empty string to allow comparison
    if (message.lastModifiedDateTime && message.lastModifiedDateTime > (cachedData.lastModifiedDateTime ?? ''))
      cachedData.lastModifiedDateTime = message.lastModifiedDateTime;
  }

  public async deleteMessage(chatId: string, message: ChatMessage) {
    const cachedData = await this.cache.getValue(chatId);
    // for now we're ignoring the case where we didn't find anything in the cache for the given chatId as there's nothing to delete.
    if (cachedData) {
      const spliceIndex = cachedData.value.findIndex(m => m.id === message.id);
      cachedData.value.splice(spliceIndex, 1);
      await this.cache.putValue(chatId, cachedData);
    }
  }
}
