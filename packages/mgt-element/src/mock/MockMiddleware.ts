/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';

/**
 * Implements Middleware for the Mock Client to escape
 * the graph url from the request
 *
 * @class MockMiddleware
 * @implements {Middleware}
 */
export class MockMiddleware implements Middleware {
  /**
   * @private
   * A member to hold next middleware in the middleware chain
   */
  private _nextMiddleware: Middleware;

  private static _baseUrl: string;

  public async execute(context: Context): Promise<void> {
    try {
      const baseUrl = await MockMiddleware.getBaseUrl();
      context.request = baseUrl + encodeURIComponent(context.request as string);
    } catch (error) {
      // ignore error
    }
    return await this._nextMiddleware.execute(context);
  }
  /**
   * Handles setting of next middleware
   *
   * @param {Middleware} next
   * @memberof SdkVersionMiddleware
   */
  public setNext(next: Middleware): void {
    this._nextMiddleware = next;
  }

  /**
   * Gets the base url for the mock graph.
   *
   * @static
   * @return {string} the base url for the mock graph to use.
   * @memberof MockMiddleware
   */
  public static async getBaseUrl() {
    if (!this._baseUrl) {
      this._baseUrl = 'https://graphexplorer.microsoft.com/api/proxy?url=';
    }

    return await Promise.resolve(this._baseUrl);
  }
}
