/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { SessionCache, storageAvailable } from '../utils/SessionCache';

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

  private static _cache: SessionCache;
  private static get _sessionCache(): SessionCache {
    if (!this._cache && storageAvailable('sessionStorage')) {
      this._cache = new SessionCache();
    }
    return this._cache;
  }

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
   * Gets the base url for the mock graph, either from the session cache or from the endpoint service
   *
   * @static
   * @return {string} the base url for the mock graph to use.
   * @memberof MockMiddleware
   */
  public static async getBaseUrl() {
    if (!this._baseUrl) {
      const sessionEndpoint = this._sessionCache?.getItem('endpointURL');
      if (sessionEndpoint) {
        this._baseUrl = sessionEndpoint;
      } else {
        try {
          // get the url we should be using from the endpoint service
          const response = await fetch('https://cdn.graph.office.net/en-us/graph/api/proxy/endpoint');
          const base: unknown = await response.json();
          if (typeof base !== 'string') {
            MockMiddleware.setBaseFallbackUrl();
          } else {
            this._baseUrl = base + '?url=';
          }
        } catch {
          // fallback to hardcoded value
          MockMiddleware.setBaseFallbackUrl();
        }
        this._sessionCache?.setItem('endpointURL', this._baseUrl);
      }
    }

    return this._baseUrl;
  }

  private static setBaseFallbackUrl() {
    this._baseUrl = 'https://proxy.apisandbox.msdn.microsoft.com/svc?url=';
  }
}
