/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AuthenticationHandler,
  Client,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import { MgtBaseComponent } from '../components/baseComponent';
import { Graph } from '../Graph';
import { chainMiddleware } from '../utils/GraphHelpers';

import { MockProvider } from './MockProvider';
import { MockMiddleware } from './MockMiddleware';

/**
 * MockGraph Instance
 *
 * @export
 * @class MockGraph
 * @extends {Graph}
 */
export class MockGraph extends Graph {
  /**
   * Creates a new MockGraph instance. Use this static method instead of the constructor.
   *
   * @static
   * @param {MockProvider} provider
   * @return {*}  {Promise<MockGraph>}
   * @memberof MockGraph
   */
  public static async create(provider: MockProvider): Promise<MockGraph> {
    const middleware: Middleware[] = [
      new AuthenticationHandler(provider),
      new RetryHandler(new RetryHandlerOptions()),
      new TelemetryHandler(),
      new MockMiddleware(),
      new HTTPMessageHandler()
    ];

    return new MockGraph(
      Client.initWithMiddleware({
        middleware: chainMiddleware(...middleware),
        customHosts: new Set<string>([new URL(await MockMiddleware.getBaseUrl()).hostname])
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
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  public forComponent(component: MgtBaseComponent): MockGraph {
    // The purpose of the forComponent pattern is to update the headers of any outgoing Graph requests.
    // The MockGraph isn't making real Graph requests, so we can simply no-op and return the same instance.
    return this;
  }
}
