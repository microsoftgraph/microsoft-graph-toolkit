/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationHandlerOptions, Middleware } from '@microsoft/microsoft-graph-client';
import { GraphEndpoint, MICROSOFT_GRAPH_ENDPOINTS } from '../IGraph';
import { Providers } from '../providers/Providers';

/**
 * creates an AuthenticationHandlerOptions from scopes array that
 * can be used in the Graph sdk middleware chain
 *
 * @export
 * @param {...string[]} scopes
 * @returns
 */
export const prepScopes = (...scopes: string[]) => {
  const authProviderOptions = {
    scopes
  };

  if (!Providers.globalProvider.isIncrementalConsentDisabled) {
    return [new AuthenticationHandlerOptions(undefined, authProviderOptions)];
  } else {
    return [];
  }
};

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

/**
 * Helper method to validate a base URL string
 *
 * @param url a URL string
 * @returns GraphEndpoint
 */
export const validateBaseURL = (url: string): GraphEndpoint => {
  try {
    const urlObj = new URL(url);
    const originAsEndpoint = urlObj.origin as GraphEndpoint;
    if (MICROSOFT_GRAPH_ENDPOINTS.has(originAsEndpoint)) {
      return originAsEndpoint;
    }
  } catch (error) {
    return;
  }
};
