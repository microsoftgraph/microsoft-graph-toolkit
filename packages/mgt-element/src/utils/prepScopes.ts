/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationHandlerOptions } from '@microsoft/microsoft-graph-client';
import { Providers } from '../providers/Providers';

/**
 * creates an AuthenticationHandlerOptions from scopes array that
 * can be used in the Graph sdk middleware chain
 *
 * @export
 * @param {...string[]} scopes
 * @returns
 */

export const prepScopes = (scopes: string[], provider = Providers.globalProvider) => {
  const additionalScopes = provider.needsAdditionalScopes(scopes);
  const authProviderOptions = {
    scopes: additionalScopes
  };

  if (!provider.isIncrementalConsentDisabled) {
    return [new AuthenticationHandlerOptions(undefined, authProviderOptions)];
  } else {
    return [];
  }
};
