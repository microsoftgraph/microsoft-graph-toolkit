/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { getRequestHeader, setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/middleware/MiddlewareUtil';
import { PACKAGE_VERSION } from './version';

/**
 * Implements Middleware for the Graph sdk to inject
 * the toolkit version in the SdkVersion header
 *
 * @class SdkVersionMiddleware
 * @implements {Middleware}
 */
export class SdkVersionMiddleware implements Middleware {
  /**
   * @private
   * A member to hold next middleware in the middleware chain
   */
  private nextMiddleware: Middleware;

  // tslint:disable-next-line: completed-docs
  public async execute(context: Context): Promise<void> {
    try {
      let sdkVersionValue: string = `mgt/${PACKAGE_VERSION}`;

      sdkVersionValue += ', ' + getRequestHeader(context.request, context.options, 'SdkVersion');

      setRequestHeader(context.request, context.options, 'SdkVersion', sdkVersionValue);
      return await this.nextMiddleware.execute(context);
    } catch (error) {
      throw error;
    }
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
