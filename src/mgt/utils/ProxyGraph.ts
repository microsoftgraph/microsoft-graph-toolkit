/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  Client,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import { chainMiddleware, Graph, SdkVersionMiddleware } from '../../mgt-core';
import { PACKAGE_VERSION } from '../../version';
import { CustomHeaderMiddleware } from './CustomHeaderMiddleware';

/**
 * ProxyGraph Instance
 *
 * @export
 * @class ProxyGraph
 * @extends {Graph}
 */
export class ProxyGraph extends Graph {
  constructor(baseUrl: string, getCustomHeaders: () => Promise<object>) {
    const middleware: Middleware[] = [
      new RetryHandler(new RetryHandlerOptions()),
      new TelemetryHandler(),
      new SdkVersionMiddleware(PACKAGE_VERSION),
      new CustomHeaderMiddleware(getCustomHeaders),
      new HTTPMessageHandler()
    ];

    super(
      Client.initWithMiddleware({
        baseUrl,
        middleware: chainMiddleware(...middleware)
      })
    );
  }
}
