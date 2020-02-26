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
import { Graph } from './Graph';
import { CustomHeaderMiddleware } from './utils/CustomHeaderMiddleware';
import { chainMiddleware } from './utils/GraphHelpers';
import { SdkVersionMiddleware } from './utils/SdkVersionMiddleware';
import { PACKAGE_VERSION } from './utils/version';

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
