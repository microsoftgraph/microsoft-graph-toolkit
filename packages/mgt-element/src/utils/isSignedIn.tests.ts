/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { assert, expect } from '@open-wc/testing';
import { isSignedIn } from './isSignedIn';
import { MockProvider } from '../mock/MockProvider';
import { Providers } from '../providers/Providers';
import { ProviderState } from '../providers/IProvider';

describe('signedInState', () => {
  beforeEach(() => {
    Providers.globalProvider = new MockProvider(true);
  });

  it('should be false', () => expect(isSignedIn()).to.be.false);

  it('should be true', () => {
    const provider = Providers.globalProvider;
    provider.setState(ProviderState.SignedIn);
    assert(isSignedIn());
  });
});
