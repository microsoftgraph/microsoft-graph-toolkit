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
import { BaseGraph } from './BaseGraph';
import { MgtBaseComponent } from './components/baseComponent';
import { IProvider } from './providers/IProvider';
import { SdkVersionMiddleware } from './utils/SdkVersionMiddleware';
import { PACKAGE_VERSION } from './utils/version';

/**
 * Creates async methods for requesting data from the Graph
 *
 * @export
 * @class Graph
 */
export class Graph extends BaseGraph {
  constructor(providerOrClient: IProvider | Client) {
    super();

    const client = providerOrClient as Client;
    const provider = providerOrClient as IProvider;

    if (client.api) {
      this.client = client;
    } else if (provider) {
      const middleware: Middleware[] = [
        new AuthenticationHandler(provider),
        new RetryHandler(new RetryHandlerOptions()),
        new TelemetryHandler(),
        new SdkVersionMiddleware(PACKAGE_VERSION),
        new HTTPMessageHandler()
      ];

      this.client = Client.initWithMiddleware({
        middleware: this.chainMiddleware(...middleware)
      });
    }
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {MgtBaseComponent} component
   * @returns
   * @memberof BaseGraph
   */
  public forComponent(component: MgtBaseComponent): Graph {
    const graph = new Graph(this.client);
    graph.componentName = component.tagName.toLowerCase();
    return graph;
  }
}
