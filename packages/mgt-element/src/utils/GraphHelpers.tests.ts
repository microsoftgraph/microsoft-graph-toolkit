/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-unused-expressions */
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
import { fake } from 'sinon';
import { AuthenticationHandlerOptions, Middleware } from '@microsoft/microsoft-graph-client';
import { MockProvider } from '../mock/MockProvider';
import { Providers } from '../providers/Providers';
import { chainMiddleware } from './chainMiddleware';
import { prepScopes } from './prepScopes';
import { validateBaseURL } from './validateBaseURL';

describe('GraphHelpers - prepScopes', () => {
  it('should return an empty array when incremental consent is disabled', async () => {
    const scopes = ['scope1', 'scope2'];
    Providers.globalProvider = new MockProvider(true);
    Providers.globalProvider.isIncrementalConsentDisabled = true;
    // eql for loose equality
    await expect(prepScopes(scopes)).to.eql([]);
  });
  it('should return an array of AuthenticationHandlerOptions when incremental consent is enabled with only the first scope in the list', async () => {
    const scopes = ['scope1', 'scope2'];
    Providers.globalProvider = new MockProvider(true);
    Providers.globalProvider.isIncrementalConsentDisabled = false;
    await expect(prepScopes(scopes)).to.eql([new AuthenticationHandlerOptions(undefined, { scopes: ['scope1'] })]);
  });
});

describe('GraphHelpers - chainMiddleware', () => {
  it('should return the first middleware when only one is passed', async () => {
    const middleware: Middleware[] = [{ execute: fake(), setNext: fake() }];
    const result = chainMiddleware(...middleware);
    await expect(result).to.equal(middleware[0]);
  });

  it('should return undefined when the middleware array is empty', () => {
    const middleware: Middleware[] = [];
    const result = chainMiddleware(...middleware);
    expect(result).to.be.undefined;
  });
  it('should now throw when the middleware array is undefined', () => {
    let error: string;
    try {
      const result = chainMiddleware(undefined);
      expect(result).to.be.undefined;
    } catch (e) {
      error = 'thrown and caught';
    }
    expect(error).to.be.undefined;
  });
});

describe('GraphHelpers - validateBaseUrl', () => {
  it('should return as a valid Url', async () => {
    const validUrls = [
      'https://graph.microsoft.com',
      'https://graph.microsoft.us',
      'https://dod-graph.microsoft.us',
      'https://graph.microsoft.de',
      'https://microsoftgraph.chinacloudapi.cn'
    ];
    for (const url of validUrls) {
      await expect(validateBaseURL(url)).to.equal(url);
    }
  });
  it('should return undefeined for invalid Url', () => {
    const validUrls = ['https://graph.microsoft.net', 'https://random.us', 'https://nope.cn'];
    for (const url of validUrls) {
      expect(validateBaseURL(url)).to.be.undefined;
    }
  });

  it('should return undefined for when supplied a %p which is not a well formed url', () => {
    const testValues = ['not a url', 'graph.microsoft.com'];
    for (const test of testValues) {
      expect(validateBaseURL(test)).to.be.undefined;
    }
  });
});
