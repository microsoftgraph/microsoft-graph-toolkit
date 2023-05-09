/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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
