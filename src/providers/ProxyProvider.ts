/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  Client,
  Context,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import { setRequestHeader } from '@microsoft/microsoft-graph-client/lib/es/middleware/MiddlewareUtil';
import { Graph } from '../Graph';
import { IProvider, ProviderState } from '../providers/IProvider';
import { SdkVersionMiddleware } from '../utils/SdkVersionMiddleware';
/**
 * Proxy Provider access token for Microsoft Graph APIs
 *
 * @export
 * @class ProxyProvider
 * @extends {IProvider}
 */
export class ProxyProvider extends IProvider {
  /**
   * new instance of proxy graph provider
   *
   * @memberof ProxyProvider
   */
  public graph: Graph;
  constructor(graphProxyUrl: string, getCustomHeaders: () => Promise<object> = null) {
    super();
    this.graph = new ProxyGraph(graphProxyUrl, getCustomHeaders);
    this.graph.getMe().then(
      user => {
        if (user != null) {
          this.setState(ProviderState.SignedIn);
        } else {
          this.setState(ProviderState.SignedOut);
        }
      },
      err => {
        this.setState(ProviderState.SignedOut);
      }
    );
  }

  /**
   * Promise returning token
   *
   * @returns {Promise<string>}
   * @memberof ProxyProvider
   */
  public getAccessToken(): Promise<string> {
    return null;
  }
}

/**
 * ProxyGraph Instance
 *
 * @export
 * @class ProxyGraph
 * @extends {Graph}
 */
// tslint:disable-next-line: max-classes-per-file
export class ProxyGraph extends Graph {
  constructor(baseUrl: string, getCustomHeaders: () => Promise<object>) {
    super(null);

    const retryHandler = new RetryHandler(new RetryHandlerOptions());
    const telemetryHandler = new TelemetryHandler();
    const sdkVersionMiddleware = new SdkVersionMiddleware();
    const customHeaderMiddleware = new CustomHeaderMiddleware(getCustomHeaders);
    const httpMessageHandler = new HTTPMessageHandler();

    retryHandler.setNext(telemetryHandler);
    telemetryHandler.setNext(sdkVersionMiddleware);
    sdkVersionMiddleware.setNext(customHeaderMiddleware);
    customHeaderMiddleware.setNext(httpMessageHandler);

    this.client = Client.initWithMiddleware({
      baseUrl,
      middleware: retryHandler
    });
  }
}

/**
 * Custom Middleware to add custom headers when making calls
 * through the proxy provider
 *
 * @class CustomHeaderMiddleware
 * @implements {Middleware}
 */
// tslint:disable-next-line: max-classes-per-file
class CustomHeaderMiddleware implements Middleware {
  private nextMiddleware: Middleware;
  private _getCustomHeaders: () => Promise<object>;

  public constructor(getCustomHeaders: () => Promise<object>) {
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
