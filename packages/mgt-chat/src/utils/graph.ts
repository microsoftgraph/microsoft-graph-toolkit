/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, Providers } from '@microsoft/mgt-element';

export const graph = (component: string, version?: string): IGraph => {
  if (!component) throw new Error('Component name is required');
  // NOTE: for some unknown reason, the global provider is not available on first
  // load. A delay is required to ensure that it is instatiated here.
  return Providers.globalProvider?.graph.forComponent(component, version);
};
