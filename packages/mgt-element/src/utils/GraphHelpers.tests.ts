/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { it } from '@jest/globals';
import { AuthenticationHandlerOptions, Middleware } from '@microsoft/microsoft-graph-client';
import { MockProvider } from '../mock/MockProvider';
import { Providers } from '../providers/Providers';
import { chainMiddleware, prepScopes, validateBaseURL } from './GraphHelpers';

describe('GraphHelpers - prepScopes', () => {
  it('should return an empty array when incremental consent is disabled', () => {
    const scopes = ['scope1', 'scope2'];
    Providers.globalProvider = new MockProvider(true);
    Providers.globalProvider.isIncrementalConsentDisabled = true;
    expect(prepScopes(...scopes)).toEqual([]);
  });
  it('should return an array of AuthenticationHandlerOptions when incremental consent is enabled', () => {
    const scopes = ['scope1', 'scope2'];
    Providers.globalProvider = new MockProvider(true);
    Providers.globalProvider.isIncrementalConsentDisabled = false;
    expect(prepScopes(...scopes)).toEqual([new AuthenticationHandlerOptions(undefined, { scopes })]);
  });
});

describe('GraphHelpers - chainMiddleware', () => {
  it('should return the first middleware when only one is passed', () => {
    const middleware: Middleware[] = [{ execute: jest.fn(), setNext: jest.fn() }];
    const result = chainMiddleware(...middleware);
    expect(result).toEqual(middleware[0]);
  });

  it('should return undefined when the middleware array is empty', () => {
    const middleware: Middleware[] = [];
    const result = chainMiddleware(...middleware);
    expect(result).toBeUndefined();
  });
  it('should now throw when the middleware array is undefined', () => {
    let error: string;
    try {
      const result = chainMiddleware(undefined);
      expect(result).toBeUndefined();
    } catch (e) {
      error = 'thrown and caught';
    }
    expect(error).toBeUndefined();
  });
});

describe('GraphHelpers - validateBaseUrl', () => {
  it.each([
    'https://graph.microsoft.com',
    'https://graph.microsoft.us',
    'https://dod-graph.microsoft.us',
    'https://graph.microsoft.de',
    'https://microsoftgraph.chinacloudapi.cn'
  ])('should return %p as a valid base url', (graphUrl: string) => {
    expect(validateBaseURL(graphUrl)).toBe(graphUrl);
  });
  it.each(['https://graph.microsoft.net', 'https://random.us', 'https://nope.cn'])(
    'should return undefined for %p as an invalid base url',
    (graphUrl: string) => {
      expect(validateBaseURL(graphUrl)).toBeUndefined();
    }
  );
  it.each(['not a url', 'graph.microsoft.com'])(
    'should return undefined for when supplied a %p which is not a well formed url',
    (input: string) => {
      expect(validateBaseURL(input)).toBeUndefined();
    }
  );
});
