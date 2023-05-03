import { CacheItem } from '@microsoft/mgt-element';

/**
 * Object to be stored in cache representing a generic query
 */
export interface CacheResponse extends CacheItem {
  /**
   * json representing a response as string
   */
  response?: string;
}
