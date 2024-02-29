/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ProviderState } from '../providers/IProvider';
import { Providers } from '../providers/Providers';

/**
 * Checks the provider state if it's in signed in.
 *
 * @returns true if signed in, otherwise false.
 */
export const isSignedIn = (): boolean => {
  const provider = Providers.globalProvider;
  return provider && provider.state === ProviderState.SignedIn;
};
