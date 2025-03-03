/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore, CacheSchema } from '@microsoft/mgt-element';
import { Contact, Person, User } from '@microsoft/microsoft-graph-types';
import { CollectionResponse } from '@microsoft/mgt-element';
import { extractEmailAddress } from '../utils/Utils';
import { schemas } from './cacheStores';
import { IDynamicPerson } from './types';

const personTypes = ['any', 'person', 'group'] as const;
/**
 * Person Type enum
 *
 * @export
 * @enum {string}
 */
export type PersonType = (typeof personTypes)[number];
export const isPersonType = (value: unknown): value is PersonType =>
  typeof value === 'string' && personTypes.includes(value as PersonType);
export const personTypeConverter = (value: string, defaultValue: PersonType = 'any'): PersonType =>
  isPersonType(value) ? value : defaultValue;

const userTypes = ['any', 'user', 'contact'] as const;
/**
 * User Type enum
 *
 * @export
 * @enum {string}
 */
export type UserType = (typeof userTypes)[number];

export const isUserType = (value: unknown): value is UserType => {
  return typeof value === 'string' && userTypes.includes(value as UserType);
};

export const userTypeConverter = (value: string, defaultValue: UserType = 'any'): UserType =>
  isUserType(value) ? value : defaultValue;

/**
 * Object to be stored in cache representing individual people
 */
interface CachePerson extends CacheItem {
  /**
   * json representing a person stored as string
   */
  person?: string;
}

/**
 * Stores results of queries (multiple people returned)
 */
interface CachePeopleQuery extends CacheItem {
  /**
   * max number of results the query asks for
   */
  maxResults?: number;
  /**
   * list of people returned by query (might be less than max results!)
   */
  results?: string[];
}

/**
 * Defines the expiration time
 */
const getPeopleInvalidationTime = (): number => {
  return CacheService.config.people.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;
};

/**
 * Whether the people store is enabled
 */
const getIsPeopleCacheEnabled = (): boolean => CacheService.config.people.isEnabled && CacheService.config.isEnabled;

const validPeopleQueryScopes = ['People.Read', 'People.Read.All'];
const validContactQueryScopes = ['Contacts.Read', 'Contacts.ReadWrite'];

/**
 * async promise, returns all Graph people who are most relevant contacts to the signed in user.
 *
 * @param {IGraph} graph
 * @param {string} query
 * @param {number} [top=10] - number of people to return
 * @param {UserType} [personType='any'] - the type of person to search for
 * @returns {(Promise<Person[]>)}
 */
export const findPeople = async (
  graph: IGraph,
  query: string,
  top = 10,
  userType: UserType = 'any',
  filters = ''
): Promise<Person[]> => {
  const cacheKey = `${query}:${top}:${userType}`;
  let cache: CacheStore<CachePeopleQuery>;
  if (getIsPeopleCacheEnabled()) {
    const people: CacheSchema = schemas.people;
    const peopleQuery: string = schemas.people.stores.peopleQuery;
    cache = CacheService.getCache<CachePeopleQuery>(people, peopleQuery);
    const result: CachePeopleQuery = getIsPeopleCacheEnabled() ? await cache.getValue(cacheKey) : null;
    if (result && getPeopleInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(peopleStr => JSON.parse(peopleStr) as Person);
    }
  }

  let filter = "personType/class eq 'Person'";

  if (userType !== 'any') {
    if (userType === 'user') {
      filter += "and personType/subclass eq 'OrganizationUser'";
    } else {
      filter += "and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')";
    }
  }

  if (filters !== '') {
    // Adding the default people filters to the search filters
    filter += ` and  ${filters}`;
  }
  let graphResult: CollectionResponse<Person>;
  try {
    let graphRequest = graph
      .api('/me/people')
      .search('"' + query + '"')
      .top(top)
      .filter(filter)
      .middlewareOptions(prepScopes(validPeopleQueryScopes));

    if (userType !== 'contact') {
      // for any type other than Contact, user a wider search
      graphRequest = graphRequest.header('X-PeopleQuery-QuerySources', 'Mailbox,Directory');
    }

    graphResult = (await graphRequest.get()) as CollectionResponse<Person>;

    if (getIsPeopleCacheEnabled() && graphResult) {
      const item: CachePeopleQuery = { maxResults: top, results: null };
      item.results = graphResult.value.map(personStr => JSON.stringify(personStr));
      await cache.putValue(cacheKey, item);
    }
  } catch (error) {
    // intentionally empty
  }
  return graphResult?.value;
};

/**
 * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export const getPeople = async (
  graph: IGraph,
  userType: UserType = 'any',
  peopleFilters = '',
  top = 10
): Promise<Person[]> => {
  let cache: CacheStore<CachePeopleQuery>;
  const cacheKey = `${peopleFilters ? peopleFilters : `*:${userType}`}:${top}`;

  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(schemas.people, schemas.people.stores.peopleQuery);
    const cacheRes = await cache.getValue(cacheKey);

    if (cacheRes && getPeopleInvalidationTime() > Date.now() - cacheRes.timeCached) {
      return cacheRes.results.map(ppl => JSON.parse(ppl) as Person);
    }
  }

  const uri = '/me/people';
  let filter = "personType/class eq 'Person'";
  if (userType !== 'any') {
    if (userType === 'user') {
      filter += "and personType/subclass eq 'OrganizationUser'";
    } else {
      filter += "and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')";
    }
  }

  if (peopleFilters) {
    filter += ` and ${peopleFilters}`;
  }

  let people: CollectionResponse<Person>;
  try {
    let graphRequest = graph.api(uri).middlewareOptions(prepScopes(validPeopleQueryScopes)).top(top).filter(filter);

    if (userType !== 'contact') {
      // for any type other than Contact, user a wider search
      graphRequest = graphRequest.header('X-PeopleQuery-QuerySources', 'Mailbox,Directory');
    }

    people = (await graphRequest.get()) as CollectionResponse<Person>;
    if (getIsPeopleCacheEnabled() && people) {
      await cache.putValue(cacheKey, { maxResults: 10, results: people.value.map(ppl => JSON.stringify(ppl)) });
    }
  } catch (_) {
    // no-op
  }
  return people ? people.value : null;
};

/**
 * Attempts to extract the email from the IDynamicPerson properties.
 *
 * @param {IDynamicperson} entity
 */
export const getEmailFromGraphEntity = (entity: IDynamicPerson): string => {
  const person = entity as Person;
  const user = entity as User;
  const contact = entity as Contact;

  if (user?.mail) {
    return extractEmailAddress(user.mail);
  } else if (person?.scoredEmailAddresses?.length) {
    return extractEmailAddress(person.scoredEmailAddresses[0].address);
  } else if (contact?.emailAddresses?.length) {
    return extractEmailAddress(contact.emailAddresses[0].address);
  }
  return null;
};

/**
 * async promise, returns a Graph contact associated with the email provided
 *
 * @param {string} email
 * @returns {(Promise<Contact[]>)}
 * @memberof Graph
 */
export const findContactsByEmail = async (graph: IGraph, email: string): Promise<Contact[]> => {
  let cache: CacheStore<CachePerson>;
  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePerson>(schemas.people, schemas.people.stores.contacts);
    const contact = await cache.getValue(email);

    if (contact && getPeopleInvalidationTime() > Date.now() - contact.timeCached) {
      return JSON.parse(contact.person) as Contact[];
    }
  }

  const encodedEmail = encodeURIComponent(email);

  const result = (await graph
    .api('/me/contacts')
    .filter(`emailAddresses/any(a:a/address eq '${encodedEmail}')`)
    .middlewareOptions(prepScopes(validContactQueryScopes))
    .get()) as CollectionResponse<Contact>;

  if (getIsPeopleCacheEnabled() && result) {
    await cache.putValue(email, { person: JSON.stringify(result.value) });
  }

  return result ? result.value : null;
};

/**
 * async promise, returns Graph people matching the Graph query specified
 * in the resource param
 *
 * @param {IGraph} graph - the graph instance to use for making requests
 * @param {string} version - the graph version url segment to use when making requests
 * @param {string} resource - the resource segment of the graph url to be requested
 * @param {string[]} scopes - an array of scopes that are required to make the underlying graph request,
 *  if any scope provided is not currently consented then the user will be prompted for consent prior to
 *  making the graph request to load data.
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export const getPeopleFromResource = async (
  graph: IGraph,
  version: string,
  resource: string,
  scopes: string[]
): Promise<Person[]> => {
  let cache: CacheStore<CachePeopleQuery>;
  const key = `${version}${resource}`;
  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(schemas.people, schemas.people.stores.peopleQuery);
    const result: CachePeopleQuery = await cache.getValue(key);
    if (result && getPeopleInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(peopleStr => JSON.parse(peopleStr) as Person);
    }
  }

  let request = graph.api(resource).version(version);

  if (scopes?.length) {
    request = request.middlewareOptions(prepScopes(scopes));
  }

  let response = (await request.get()) as CollectionResponse<Person>;
  // get more pages if there are available
  if (response && Array.isArray(response.value) && response['@odata.nextLink']) {
    let page = response;

    while (page?.['@odata.nextLink']) {
      const nextLink = page['@odata.nextLink'] as string;
      const nextResource = nextLink.split(version)[1];
      page = (await graph.api(nextResource).version(version).get()) as CollectionResponse<Person>;
      if (page?.value?.length) {
        page.value = response.value.concat(page.value);
        response = page;
      }
    }
  }

  if (getIsPeopleCacheEnabled() && response) {
    const item = { results: null };
    if (Array.isArray(response.value)) {
      item.results = response.value.map(personStr => JSON.stringify(personStr));
    } else {
      item.results = [JSON.stringify(response)];
    }
    await cache.putValue(key, item);
  }

  return response?.value;
};
