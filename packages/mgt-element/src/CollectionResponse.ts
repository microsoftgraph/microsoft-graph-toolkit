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
  value: T[];
}
