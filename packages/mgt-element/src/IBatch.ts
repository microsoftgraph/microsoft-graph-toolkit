/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Represents a response from a request executed in a batch
 *
 * @export
 * @class BatchResponse
 */
export class BatchResponse<T = any> {
  /**
   * The index of the requests as it was added to the queue
   * Use this value if you need to sort the responses
   * in the order they were added
   *
   * @type {number}
   * @memberof BatchResponse
   */
  public index: number;

  /**
   * The id of the requests
   *
   * @type {string}
   * @memberof BatchResponse
   */
  public id: string;

  /**
   * The content of the response
   *
   * @type {*}
   * @memberof BatchResponse
   */
  public content: T;

  /**
   * The header of the response
   *
   * @type {*}
   * @memberof BatchResponse
   */
  public headers: Record<string, string>;
}

/**
 * Represents a response from a request executed in a batch
 *
 * @interface ResponseWithBodyAndStatus
 */
interface ResponseWithBodyAndStatus {
  /**
   * The body of the response
   *
   * @type {string}
   * @memberof BatchResponse
   */
  body: any;

  /**
   * The status code of the response
   *
   * @type {number}
   * @memberof BatchResponse
   */
  status: number;
}

/**
 * Wrapper for the response body of a batch request
 */
export interface BatchResponseBody {
  /**
   * Collection of responses from the batch request
   *
   * @type {BatchResponse[]}
   * @memberof BatchResponseBody
   */
  responses: (Omit<BatchResponse, 'content'> & ResponseWithBodyAndStatus)[];
}

/**
 * Represents a collection of Graph requests to be sent simultaneously.
 *
 * @export
 * @interface IBatch
 */
export interface IBatch<T = any> {
  /**
   * Get whether there are requests that have not been executed
   *
   * @readonly
   * @memberof IBatch
   */
  readonly hasRequests: boolean;

  /**
   * sets new request and scopes
   *
   * @param {string} id
   * @param {string} resource
   * @param {string[]} [scopes]
   * @param {{ [header: string]: string }} [headers]
   * @memberof IBatch
   */
  get(id: string, resource: string, scopes?: string[], headers?: Record<string, string>);

  /**
   * Execute the next set of requests.
   * This will execute up to 20 requests at a time
   *
   * @returns {Promise<Map<string, BatchResponse<T>>>}
   * @memberof IBatch
   */
  executeNext(): Promise<Map<string, BatchResponse<T>>>;

  /**
   * Execute all requests, up to 20 at a time until
   * all requests have been executed
   *
   * @returns {Promise<Map<string, BatchResponse<T>>>}
   * @memberof IBatch
   */
  executeAll(): Promise<Map<string, BatchResponse<T>>>;
}
