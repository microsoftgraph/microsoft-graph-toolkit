/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Holder type for collection responses
 *
 * @interface CollectionResponse
 * @template T
 */

export interface CollectionResponse<T> {
  /**
   * The collection of items
   */
  value?: T[];
}
