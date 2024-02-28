/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
import { signedInState } from './signedInState';
import { MockProvider } from '../mock/MockProvider';
import { Providers } from '../providers/Providers';
import { ProviderState } from '../providers/IProvider';

describe('signedInState', () => {
  it('should be false', (): void => {
    Providers.globalProvider = new MockProvider(true);
    expect(signedInState()).to.be.false;
  });

  it('should be true', () => {
    Providers.globalProvider = new MockProvider(true);
    const provider = Providers.globalProvider;
    provider.setState(ProviderState.SignedIn);
    expect(signedInState()).to.be.true;
  });
});
