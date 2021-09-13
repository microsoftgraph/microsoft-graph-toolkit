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
// tslint:disable-next-line:max-classes-per-file
export class BatchResponse {
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
  public content: any;
  /**
   * The header of the response
   * @type {*}
   * @memberof BatchResponse
   */
  public headers: string[];
}

/**
 * Represents a collection of Graph requests to be sent simultaneously.
 *
 * @export
 * @interface IBatch
 */
export interface IBatch {
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
  get(id: string, resource: string, scopes?: string[], headers?: { [header: string]: string });

  /**
   * Execute the next set of requests.
   * This will execute up to 20 requests at a time
   *
   * @returns {Promise<Map<string, BatchResponse>>}
   * @memberof IBatch
   */
  executeNext(): Promise<Map<string, BatchResponse>>;

  /**
   * Execute all requests, up to 20 at a time until
   * all requests have been executed
   *
   * @returns {Promise<Map<string, BatchResponse>>}
   * @memberof IBatch
   */
  executeAll(): Promise<Map<string, BatchResponse>>;
}
