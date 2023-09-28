/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { screen } from 'testing-library__dom';
import { fixture } from '@open-wc/testing-helpers';
import { enableFetchMocks } from 'jest-fetch-mock';
import fetchMock from 'jest-fetch-mock';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { userPhotoBatchResponse } from './__test_data/mock-responses';
import './mgt-person';
enableFetchMocks();

let person: Element;
describe('mgt-person - tests', () => {
  beforeEach(() => {
    fetchMock.resetMocks();
    fetchMock
      // Response for the setup of the MockGraph to call to the proxy service
      .once(() =>
        Promise.resolve({
          headers: {
            'Content-Type': 'application/json'
          },
          status: 200,
          body: 'https://fake-proxy.microsoft.com/'
        })
      )
      // response to the $batch request to load the user photo and user data when there is nothing in the cache
      // note that this used mockOnceIf matching the fake url supplied in the first mock.
      .mockOnceIf(/https:\/\/fake-proxy\.microsoft\.com\/*$/, () =>
        Promise.resolve({
          headers: {
            'Content-Type': 'application/json'
          },
          status: 200,
          body: JSON.stringify(userPhotoBatchResponse)
        })
      );
  });

  it('should render', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture('<mgt-person person-query="me" view="twoLines"></mgt-person>');
    const img = await screen.findAllByAltText('Photo for Megan Bowen');
    expect(img).not.toBeNull();
    expect(person).not.toBeUndefined();
  });

  it('should pop up a flyout on click', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture('<mgt-person person-query="me" view="twoLines" person-card="click"></mgt-person>');
    const img = await screen.findAllByAltText('Photo for Megan Bowen');
    expect(img).not.toBeNull();
    expect(img.length).toBe(1);
    // test that there is no flyout
    expect(screen.queryByTestId('flyout-slot')).toBeNull();

    img[0].click();

    expect(screen.queryByTestId('flyout-slot')).toBeDefined();
  });
});
