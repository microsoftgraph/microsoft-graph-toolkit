/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/src/middleware/MiddlewareUtil';

/**
 * Custom Middleware to add custom headers when making calls
 * through the proxy provider
 *
 * @class CustomHeaderMiddleware
 * @implements {Middleware}
 */
export class CustomHeaderMiddleware implements Middleware {
  private _nextMiddleware: Middleware;
  private readonly _getCustomHeaders?: () => Promise<object>;

  constructor(getCustomHeaders?: () => Promise<object>) {
    this._getCustomHeaders = getCustomHeaders;
  }

  /**
   * Execute the current middleware
   *
   * @param {Context} context
   * @returns {Promise<void>}
   * @memberof CustomHeaderMiddleware
   */
  public async execute(context: Context): Promise<void> {
    if (this._getCustomHeaders) {
      const headers = await this._getCustomHeaders();
      for (const key in headers) {
        if (Object.prototype.hasOwnProperty.call(headers, key)) {
          setRequestHeader(context.request, context.options, key, headers[key] as string);
        }
      }
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
}
