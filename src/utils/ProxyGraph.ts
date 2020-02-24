/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  Client,
  HTTPMessageHandler,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import { Graph } from '../Graph';
import { CustomHeaderMiddleware } from './CustomHeaderMiddleware';
import { SdkVersionMiddleware } from './SdkVersionMiddleware';
import { PACKAGE_VERSION } from './version';

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
    const retryHandler = new RetryHandler(new RetryHandlerOptions());
    const telemetryHandler = new TelemetryHandler();
    const sdkVersionMiddleware = new SdkVersionMiddleware(PACKAGE_VERSION);
    const customHeaderMiddleware = new CustomHeaderMiddleware(getCustomHeaders);
    const httpMessageHandler = new HTTPMessageHandler();

    retryHandler.setNext(telemetryHandler);
    telemetryHandler.setNext(sdkVersionMiddleware);
    sdkVersionMiddleware.setNext(customHeaderMiddleware);
    customHeaderMiddleware.setNext(httpMessageHandler);

    const client = Client.initWithMiddleware({
      baseUrl,
      middleware: retryHandler
    });
    super(client);
  }
}
