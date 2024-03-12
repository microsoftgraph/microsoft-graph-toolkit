/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
import { MockProvider } from '../mock/MockProvider';
import { ProviderState } from '../providers/IProvider';
import { Providers } from '../providers/Providers';
import { isSignedIn } from './isSignedIn';

describe('signedInState', () => {
  it('should change', () => {
    /* eslint-disable @typescript-eslint/no-unused-expressions */
    Providers.globalProvider = new MockProvider(false);
    expect(isSignedIn()).to.be.false;
    Providers.globalProvider.setState(ProviderState.SignedIn);
    expect(isSignedIn()).to.be.true;
  });
});
