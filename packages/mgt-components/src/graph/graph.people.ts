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
const getPeopleInvalidationTime = (): number => {
  return CacheService.config.people.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;
};

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
export const findPeople = async (
  graph: IGraph,
  query: string,
  top = 10,
  userType: UserType = UserType.any,
  filters = ''
): Promise<Person[]> => {
  const scopes = 'people.read';

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
  let graphResult: CollectionResponse<Person>;
  try {
    let graphRequest = graph
      .api('/me/people')
      .search('"' + query + '"')
      .top(top)
      .filter(filter)
      .middlewareOptions(prepScopes(scopes));

    if (userType !== UserType.contact) {
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
  userType: UserType = UserType.any,
  peopleFilters = '',
  top = 10
): Promise<Person[]> => {
  const scopes = 'people.read';

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

  let people: CollectionResponse<Person>;
  try {
    let graphRequest = graph.api(uri).middlewareOptions(prepScopes(scopes)).top(top).filter(filter);

    if (userType !== UserType.contact) {
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

  if (user.mail) {
    return extractEmailAddress(user.mail);
  } else if (person.scoredEmailAddresses?.length) {
    return extractEmailAddress(person.scoredEmailAddresses[0].address);
  } else if (contact.emailAddresses?.length) {
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
  const scopes = 'contacts.read';
  let cache: CacheStore<CachePerson>;
  if (getIsPeopleCacheEnabled()) {
    cache = CacheService.getCache<CachePerson>(schemas.people, schemas.people.stores.contacts);
    const contact = await cache.getValue(email);

    if (contact && getPeopleInvalidationTime() > Date.now() - contact.timeCached) {
      return JSON.parse(contact.person) as Contact[];
    }
  }

  const encodedEmail = `${email.replace(/#/g, '%2523')}`;

  const result = (await graph
    .api('/me/contacts')
    .filter(`emailAddresses/any(a:a/address eq '${encodedEmail}')`)
    .middlewareOptions(prepScopes(scopes))
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
 * @param {string} resource
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
    request = request.middlewareOptions(prepScopes(...scopes));
  }

  let response = (await request.get()) as CollectionResponse<Person>;
  // get more pages if there are available
  if (response && Array.isArray(response.value) && response['@odata.nextLink']) {
    let page = response;

    while (page?.['@odata.nextLink']) {
      const nextLink = page['@odata.nextLink'] as string;
      const nextResource = nextLink.split(version)[1];
      page = (await graph.client.api(nextResource).version(version).get()) as CollectionResponse<Person>;
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
