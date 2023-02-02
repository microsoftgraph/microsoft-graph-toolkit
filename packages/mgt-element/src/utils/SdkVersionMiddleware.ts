/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { getRequestHeader, setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/middleware/MiddlewareUtil';
import { ComponentMiddlewareOptions } from './ComponentMiddlewareOptions';
import { validateBaseURL } from './GraphHelpers';

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
  private _nextMiddleware: Middleware;
  private _packageVersion: string;
  private _providerName: string;

  constructor(packageVersion: string, providerName?: string) {
    this._packageVersion = packageVersion;
    this._providerName = providerName;
  }

  // tslint:disable-next-line: completed-docs
  public async execute(context: Context): Promise<void> {
    try {
      if (typeof context.request === 'string') {
        if (validateBaseURL(context.request)) {
          // Header parts must follow the format: 'name/version'
          const headerParts: string[] = [];

          const componentOptions = context.middlewareControl.getMiddlewareOptions(
            ComponentMiddlewareOptions
          ) as ComponentMiddlewareOptions;

          if (componentOptions) {
            const componentVersion: string = `${componentOptions.componentName}/${this._packageVersion}`;
            headerParts.push(componentVersion);
          }

          if (this._providerName) {
            const providerVersion: string = `${this._providerName}/${this._packageVersion}`;
            headerParts.push(providerVersion);
          }

          // Package version
          const packageVersion: string = `mgt/${this._packageVersion}`;
          headerParts.push(packageVersion);

          // Existing SdkVersion header value
          headerParts.push(getRequestHeader(context.request, context.options, 'SdkVersion'));

          // Join the header parts together and update the SdkVersion request header value
          const sdkVersionHeaderValue = headerParts.join(', ');
          setRequestHeader(context.request, context.options, 'SdkVersion', sdkVersionHeaderValue);
        } else {
          delete context?.options?.headers['SdkVersion'];
        }
      }
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
}
