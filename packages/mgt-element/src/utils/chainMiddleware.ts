/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Middleware } from '@microsoft/microsoft-graph-client';

/**
 * Helper method to chain Middleware when instantiating new Client
 *
 * @param {...Middleware[]} middleware
 * @returns {Middleware}
 */

export const chainMiddleware = (...middleware: Middleware[]): Middleware => {
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
};
