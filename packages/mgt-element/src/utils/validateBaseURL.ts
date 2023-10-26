/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphEndpoint, MICROSOFT_GRAPH_ENDPOINTS } from '../IGraph';

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
