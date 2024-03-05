/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { assert } from '@open-wc/testing';
import { isSignedIn } from './isSignedIn';
import { MockProvider } from '../mock/MockProvider';
import { Providers } from '../providers/Providers';

describe('signedInState', () => {
  it('should be false', () => {
    Providers.globalProvider = new MockProvider(false);
    assert(!isSignedIn());
  });

  it('should be true', () => {
    Providers.globalProvider = new MockProvider(true);
    assert(isSignedIn);
  });
});
