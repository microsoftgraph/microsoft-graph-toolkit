/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Represents a request to be executed in a batch
 *
 * @class BatchRequest
 */

export class BatchRequest {
  /**
   * url used in request
   *
   * @type {string}
   * @memberof BatchRequest
   */
  public resource: string;

  /**
   * method passed to be requested
   *
   * @type {string}
   * @memberof BatchRequest
   */
  public method: string;

  /**
   * The index of the requests as it was added to the queue
   * Use this value if you need to sort the responses
   * in the order they were added
   *
   * @type {number}
   * @memberof BatchRequest
   */
  public index: number;

  /**
   * The headers of the request
   *
   * @type {{[headerName: string]: string}}
   * @memberof BatchRequest
   */
  public headers: Record<string, string>;

  /**
   * The id of the requests
   *
   * @type {string}
   * @memberof BatchRequest
   */
  public id: string;

  constructor(index: number, id: string, resource: string, method: string) {
    this.resource = resource.startsWith('/') ? resource : `/${resource}`;
    this.method = method;
    this.index = index;
    this.id = id;
  }
}
