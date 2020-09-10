/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AuthenticationHandler,
  Client,
  Context,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import { MgtBaseComponent } from '@microsoft/mgt-element';
import { Graph } from '../Graph';
import { chainMiddleware } from '../utils/GraphHelpers';
import { MockProvider } from './MockProvider';

/**
 * The base URL for the mock endpoint
 */
const BASE_URL = 'https://proxy.apisandbox.msdn.microsoft.com/svc?url=';

/**
 * The base URL for the graph
 */
const ROOT_GRAPH_URL = 'https://graph.microsoft.com/';

/**
 * MockGraph Instance
 *
 * @export
 * @class MockGraph
 * @extends {Graph}
 */
// tslint:disable-next-line: max-classes-per-file
export class MockGraph extends Graph {
  constructor(mockProvider: MockProvider) {
    const middleware: Middleware[] = [
      new AuthenticationHandler(mockProvider),
      new RetryHandler(new RetryHandlerOptions()),
      new TelemetryHandler(),
      new MockMiddleware(),
      new HTTPMessageHandler()
    ];

    super(
      Client.initWithMiddleware({
        baseUrl: BASE_URL + ROOT_GRAPH_URL,
        middleware: chainMiddleware(...middleware)
      })
    );
  }

  /**
   * Returns an instance of the Graph in the context of the provided component.
   *
   * @param {MgtBaseComponent} component
   * @returns
   * @memberof Graph
   */
  public forComponent(component: MgtBaseComponent): MockGraph {
    // The purpose of the forComponent pattern is to update the headers of any outgoing Graph requests.
    // The MockGraph isn't making real Graph requests, so we can simply no-op and return the same instance.
    return this;
  }
}

/**
 * Implements Middleware for the Mock Client to escape
 * the graph url from the request
 *
 * @class MockMiddleware
 * @implements {Middleware}
 */
// tslint:disable-next-line: max-classes-per-file
class MockMiddleware implements Middleware {
  /**
   * @private
   * A member to hold next middleware in the middleware chain
   */
  private _nextMiddleware: Middleware;

  // tslint:disable-next-line: completed-docs
  public async execute(context: Context): Promise<void> {
    try {
      const url = context.request as string;
      const baseLength = BASE_URL.length;
      context.request = url.substring(0, baseLength) + escape(url.substring(baseLength));
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
