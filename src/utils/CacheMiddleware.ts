/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Context, Middleware } from '@microsoft/microsoft-graph-client';
import { request } from 'http';

export class CacheMiddleware implements Middleware {
  // tslint:disable-next-line: completed-docs
  public async execute(context: Context): Promise<void> {
    try {
      if (context.options.method === 'GET' && (context.request as string).endsWith('/photo/$value')) {
        const cache = await caches.open('mgt-graph-calls');
        const cachedResponse = await cache.match(context.request);

        if (cachedResponse) {
          context.response = cachedResponse;
          return;
        }

        context.response = await fetch(context.request, context.options);
        await cache.put(context.request, context.response.clone());
      } else {
        context.response = await fetch(context.request, context.options);
      }
    } catch (error) {
      throw error;
    }
  }
}
