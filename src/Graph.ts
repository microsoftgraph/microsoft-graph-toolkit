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

/**
 * Creates async methods for requesting data from the Graph
 *
 * @export
 * @class Graph
 */
export class Graph extends BaseGraph {
  constructor(provider: IProvider, component?: MgtBaseComponent) {
    super(provider);

    if (provider) {
      const middleware: Middleware[] = [
        new AuthenticationHandler(provider),
        new RetryHandler(new RetryHandlerOptions()),
        new TelemetryHandler(),
        new SdkVersionMiddleware(component),
        new HTTPMessageHandler()
      ];

      this.client = Client.initWithMiddleware({
        middleware: this.chainMiddleware(...middleware)
      });
    }
  }

  /**
   * Returns an instance of the Graph in the context of the provided component.
   *
   * @param {MgtBaseComponent} component
   * @returns
   * @memberof Graph
   */
  public forComponent(component: MgtBaseComponent): Graph {
    return new Graph(this._provider, component);
  }

  private chainMiddleware(...middleware: Middleware[]): Middleware {
    const rootMiddleware = middleware[0];
    let current = rootMiddleware;
    for (let i = 1; i < middleware.length; ++i) {
      const next = middleware[i];
      if (current.setNext) {
        current.setNext(next);
      }
      current = next;
    }
    return rootMiddleware;
  }
}
