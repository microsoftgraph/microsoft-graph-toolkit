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

  it('should render with initials when given name and surname are supplied', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Brian',
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('BH');
    expect(initials).toBeDefined();
  });

  it('should render with initials when given name and surname are null', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('FH');
    expect(initials).toBeDefined();
  });

  it('should render with first initial when only given name is supplied', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Frank',
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('F');
    expect(initials).toBeDefined();
  });

  it('should render with first initial when only given name is populated and surname is an empty string', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Frank',
        surname: '',
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('F');
    expect(initials).toBeDefined();
  });

  it('should render with last initial when only surname is supplied', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('H');
    expect(initials).toBeDefined();
  });
  it('should render with last initial when only surname is populated and given name is an empty string', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: '',
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('H');
    expect(initials).toBeDefined();
  });

  it('should render with one initial when only displayName of one word is supplied', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('F');
    expect(initials).toBeDefined();
  });

  it('should render with two initial when only displayName of more than two words is supplied', async () => {
    Providers.globalProvider = new MockProvider(true);
    person = await fixture(
      `<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank van Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    expect(person).not.toBeUndefined();
    const initials = await screen.findByText('FV');
    expect(initials).toBeDefined();
  });
});
