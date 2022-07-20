/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { Contact, Person, User } from '@microsoft/microsoft-graph-types';
import { extractEmailAddress } from '../utils/Utils';
import { schemas } from './cacheStores';
import { getUsersForUserIds } from './graph.user';
import { IDynamicPerson } from './types';

/**
 * Person Type enum
 *
 * @export
 * @enum {number}
 */
export enum PersonType {
  /**
   * Any type
   */
  any = 0,

  /**
   * A Person such as User or Contact
   */
  person = 'person',

  /**
   * A group
   */
  group = 'group'
}

/**
 * User Type enum
 *
 * @export
 * @enum {number}
 */
export enum UserType {
  /**
   * Any user or contact
   */
  any = 'any',

  /**
   * An organization User
   */
  user = 'user',

  /**
   * An implicit or personal contact
   */
  contact = 'contact'
}

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
const getPeopleInvalidationTime = (): number =>
  CacheService.config.people.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the people store is enabled
 */
const getIsPeopleCacheEnabled = (): boolean => CacheService.config.people.isEnabled && CacheService.config.isEnabled;

/**
 * async promise, returns all Graph people who are most relevant contacts to the signed in user.
 *
 * @param {string} query
 * @param {number} [top=10] - number of people to return
 * @param {PersonType} [personType=PersonType.person] - the type of person to search for
 * @returns {(Promise<Person[]>)}
 */
export async function findPeople(
  graph: IGraph,
  query: string,
  top: number = 10,
  userType: UserType = UserType.any,
  filters: string = ''
): Promise<Person[]> {
  const scopes = 'people.read';

  let cache: CacheStore<CachePeopleQuery>;
  let cacheKey = `${query}:${top}:${userType}`;

  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(schemas.people, schemas.people.stores.peopleQuery);
    const result: CachePeopleQuery = getIsPeopleCacheEnabled() ? await cache.getValue(cacheKey) : null;
    if (result && getPeopleInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(peopleStr => JSON.parse(peopleStr));
    }
  }

  let filter = "personType/class eq 'Person'";

  if (userType !== UserType.any) {
    if (userType === UserType.user) {
      filter += "and personType/subclass eq 'OrganizationUser'";
    } else {
      filter += "and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')";
    }
  }

  if (filters !== '') {
    // Adding the default people filters to the search filters
    filter += `${filter} and ${filters}`;
  }

  let graphResult;
  try {
    graphResult = await graph
      .api('/me/people')
      .search('"' + query + '"')
      .top(top)
      .filter(filter)
      .middlewareOptions(prepScopes(scopes))
      .get();

    if (getIsPeopleCacheEnabled() && graphResult) {
      const item = { maxResults: top, results: null };
      item.results = graphResult.value.map(personStr => JSON.stringify(personStr));
      cache.putValue(cacheKey, item);
    }
  } catch (error) {}
  return graphResult ? graphResult.value : null;
}

/**
 * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeople(
  graph: IGraph,
  userType: UserType = UserType.any,
  peopleFilters: string = ''
): Promise<Person[]> {
  const scopes = 'people.read';

  let cache: CacheStore<CachePeopleQuery>;
  let cacheKey = peopleFilters ? peopleFilters : `*:${userType}`;

  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(schemas.people, schemas.people.stores.peopleQuery);
    const cacheRes = await cache.getValue(cacheKey);

    if (cacheRes && getPeopleInvalidationTime() > Date.now() - cacheRes.timeCached) {
      return cacheRes.results.map(ppl => JSON.parse(ppl));
    }
  }

  const uri = '/me/people';
  let filter = "personType/class eq 'Person'";
  if (userType !== UserType.any) {
    if (userType === UserType.user) {
      filter += "and personType/subclass eq 'OrganizationUser'";
    } else {
      filter += "and (personType/subclass eq 'ImplicitContact' or personType/subclass eq 'PersonalContact')";
    }
  }

  if (peopleFilters) {
    filter += ` and ${peopleFilters}`;
  }

  let people;
  try {
    people = await graph.api(uri).middlewareOptions(prepScopes(scopes)).filter(filter).get();
    if (getIsPeopleCacheEnabled() && people) {
      cache.putValue(cacheKey, { maxResults: 10, results: people.value.map(ppl => JSON.stringify(ppl)) });
    }
  } catch (_) {}
  return people ? people.value : null;
}

/**
 * returns a promise that resolves after specified time
 * @param time in milliseconds
 */
export function getEmailFromGraphEntity(entity: IDynamicPerson): string {
  const person = entity as Person;
  const user = entity as User;
  const contact = entity as Contact;

  if (user.mail) {
    return extractEmailAddress(user.mail);
  } else if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
    return extractEmailAddress(person.scoredEmailAddresses[0].address);
  } else if (contact.emailAddresses && contact.emailAddresses.length) {
    return extractEmailAddress(contact.emailAddresses[0].address);
  }
  return null;
}

/**
 * async promise, returns a Graph contact associated with the email provided
 *
 * @param {string} email
 * @returns {(Promise<Contact[]>)}
 * @memberof Graph
 */
export async function findContactsByEmail(graph: IGraph, email: string): Promise<Contact[]> {
  const scopes = 'contacts.read';
  let cache: CacheStore<CachePerson>;
  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePerson>(schemas.people, schemas.people.stores.contacts);
    const contact = await cache.getValue(email);

    if (contact && getPeopleInvalidationTime() > Date.now() - contact.timeCached) {
      return JSON.parse(contact.person);
    }
  }

  let encodedEmail = `${email.replace(/#/g, '%2523')}`;

  const result = await graph
    .api('/me/contacts')
    .filter(`emailAddresses/any(a:a/address eq '${encodedEmail}')`)
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (getIsPeopleCacheEnabled() && result) {
    cache.putValue(email, { person: JSON.stringify(result.value) });
  }

  return result ? result.value : null;
}

/**
 * async promise, returns Graph people matching the Graph query specified
 * in the resource param
 *
 * @param {string} resource
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeopleFromResource(
  graph: IGraph,
  version: string,
  resource: string,
  scopes: string[]
): Promise<Person[]> {
  let cache: CacheStore<CachePeopleQuery>;
  const key = `${version}${resource}`;
  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(schemas.people, schemas.people.stores.peopleQuery);
    const result: CachePeopleQuery = await cache.getValue(key);
    if (result && getPeopleInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(peopleStr => JSON.parse(peopleStr));
    }
  }

  let request = graph.api(resource).version(version);

  if (scopes && scopes.length) {
    request = request.middlewareOptions(prepScopes(...scopes));
  }

  let response = await request.get();
  // get more pages if there are available
  if (response && Array.isArray(response.value) && response['@odata.nextLink']) {
    let pageCount = 1;
    let page = response;

    while (page && page['@odata.nextLink']) {
      pageCount++;
      const nextResource = page['@odata.nextLink'].split(version)[1];
      page = await graph.client.api(nextResource).version(version).get();
      if (page && page.value && page.value.length) {
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
    cache.putValue(key, item);
  }

  if (response) {
    return Array.isArray(response.value) ? response.value : [response];
  } else {
    return null;
  }
}
