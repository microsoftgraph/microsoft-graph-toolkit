/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Represents organization for a cache
 *
 * @export
 * @interface CacheSchema
 */

export interface CacheSchema {
  /**
   * version number of cache, useful for upgrading
   *
   * @type {number}
   * @memberof CacheSchema
   */
  version: number;
  /**
   * name of the cache
   *
   * @type {string}
   * @memberof CacheSchema
   */
  name: string;
  /**
   * list of stores in the cache
   *
   * @type {{ [name: string]: CacheSchemaStore }}
   * @memberof CacheSchema
   */
  stores: Record<string, string>;
}
