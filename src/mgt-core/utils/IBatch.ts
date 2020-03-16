/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Method to reduce repetitive requests to the Graph
 *
 * @export
 * @interface IBatch
 */
export interface IBatch {
  /**
   * sets new request and scopes
   *
   * @param {string} id
   * @param {string} resource
   * @param {string[]} [scopes]
   * @memberof IBatch
   */
  get(id: string, resource: string, scopes?: string[]): void;

  /**
   * Promise to handle Graph request response
   *
   * @returns {Promise<any>}
   * @memberof IBatch
   */
  execute(): Promise<any>;
}
