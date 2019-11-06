/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Method to reduce repetitive requests to the Graph utilized in Batch
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
  constructor(resource: string, method: string) {
    if (resource.charAt(0) !== '/') {
      resource = '/' + resource;
    }
    this.resource = resource;
    this.method = method;
  }
}
