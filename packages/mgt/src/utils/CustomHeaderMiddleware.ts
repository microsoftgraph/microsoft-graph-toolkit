/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/middleware/MiddlewareUtil';

/**
 * Custom Middleware to add custom headers when making calls
 * through the proxy provider
 *
 * @class CustomHeaderMiddleware
 * @implements {Middleware}
 */
// tslint:disable-next-line: max-classes-per-file
export class CustomHeaderMiddleware implements Middleware {
  private nextMiddleware: Middleware;
  private _getCustomHeaders: () => Promise<object>;

  constructor(getCustomHeaders: () => Promise<object>) {
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
        if (headers.hasOwnProperty(key)) {
          setRequestHeader(context.request, context.options, key, headers[key]);
        }
      }
    }
    return await this.nextMiddleware.execute(context);
  }

  /**
   * Handles setting of next middleware
   *
   * @param {Middleware} next
   * @memberof SdkVersionMiddleware
   */
  public setNext(next: Middleware): void {
    this.nextMiddleware = next;
  }
}
