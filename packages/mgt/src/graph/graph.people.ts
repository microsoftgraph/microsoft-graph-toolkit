/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Contact, Person, User } from '@microsoft/microsoft-graph-types';
import { CacheItem, CacheSchema, CacheService, CacheStore } from '../utils/Cache';
import { prepScopes } from '../utils/GraphHelpers';
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
 * Definition of cache structure
 */
const cacheSchema: CacheSchema = {
  name: 'people',
  stores: {
    contacts: {},
    groupPeople: {},
    peopleQuery: {}
  },
  version: 1
};

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
 * Stores people from a group
 */
interface CacheGroupPeople extends CacheItem {
  /**
   * list of people returned by query (might be less than max results!)
   */
  people?: string[];
}

/**
 * Defines the expiration time
 */
const getPeopleInvalidationTime = (): number =>
  CacheService.config.people.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the people store is enabled
 */
const peopleCacheEnabled = (): boolean => CacheService.config.people.isEnabled && CacheService.config.isEnabled;

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
  personType: PersonType = PersonType.person
): Promise<Person[]> {
  const scopes = 'people.read';

  let cache: CacheStore<CachePeopleQuery>;
  const item = { maxResults: top, results: null };

  if (peopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(cacheSchema, 'peopleQuery');
    const result: CachePeopleQuery = peopleCacheEnabled() ? await cache.getValue(query) : null;
    if (result && getPeopleInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(peopleStr => JSON.parse(peopleStr));
    }
  }
  let filterQuery = '';

  if (personType !== PersonType.any) {
    // converts personType to capitalized case
    const personTypeString =
      personType
        .toString()
        .charAt(0)
        .toUpperCase() + personType.toString().slice(1);
    filterQuery = `personType/class eq '${personTypeString}'`;
  }

  const graphResult = await graph
    .api('/me/people')
    .search('"' + query + '"')
    .top(top)
    .filter(filterQuery)
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (peopleCacheEnabled() && graphResult) {
    item.results = graphResult.value.map(personStr => JSON.stringify(personStr));
    cache.putValue(query, item);
  }
  return graphResult ? graphResult.value : null;
}

/**
 * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeople(graph: IGraph): Promise<Person[]> {
  const scopes = 'people.read';

  let cache: CacheStore<CachePeopleQuery>;
  if (peopleCacheEnabled()) {
    cache = CacheService.getCache<CachePeopleQuery>(cacheSchema, 'peopleQuery');
    // not a great way to do this, don't know a better way
    const cacheRes = await cache.getValue('*');

    if (cacheRes && getPeopleInvalidationTime() > Date.now() - cacheRes.timeCached) {
      return cacheRes.results.map(ppl => JSON.parse(ppl));
    }
  }

  const uri = '/me/people';
  const people = await graph
    .api(uri)
    .middlewareOptions(prepScopes(scopes))
    .filter("personType/class eq 'Person'")
    .get();
  if (peopleCacheEnabled() && people) {
    cache.putValue('*', { maxResults: 10, results: people.value.map(ppl => JSON.stringify(ppl)) });
  }
  return people ? people.value : null;
}

/**
 * async promise to the Graph for People, defined by a group id
 *
 * @param {string} groupId
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeopleFromGroup(graph: IGraph, groupId: string): Promise<Person[]> {
  const scopes = 'people.read';
  let cache: CacheStore<CacheGroupPeople>;

  if (peopleCacheEnabled()) {
    cache = CacheService.getCache<CacheGroupPeople>(cacheSchema, 'groupPeople');
    const peopleItem = await cache.getValue(groupId);
    if (peopleItem && getPeopleInvalidationTime() > Date.now() - peopleItem.timeCached) {
      return peopleItem.people.map(peopleStr => JSON.parse(peopleStr));
    }
  }

  const uri = `/groups/${groupId}/members`;
  const people = await graph
    .api(uri)
    .middlewareOptions(prepScopes(scopes))
    .get();
  if (peopleCacheEnabled) {
    cache.putValue(groupId, { people: people.value.map(ppl => JSON.stringify(ppl)) });
  }
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
    return user.mail;
  } else if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
    return person.scoredEmailAddresses[0].address;
  } else if (contact.emailAddresses && contact.emailAddresses.length) {
    return contact.emailAddresses[0].address;
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
export async function findContactByEmail(graph: IGraph, email: string): Promise<Contact[]> {
  const scopes = 'contacts.read';
  let cache: CacheStore<CachePerson>;
  if (peopleCacheEnabled()) {
    cache = CacheService.getCache<CachePerson>(cacheSchema, 'contacts');
    const contact = await cache.getValue(email);

    if (contact && getPeopleInvalidationTime() > Date.now() - contact.timeCached) {
      return JSON.parse(contact.person);
    }
  }

  const result = await graph
    .api('/me/contacts')
    .filter(`emailAddresses/any(a:a/address eq '${email}')`)
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (peopleCacheEnabled() && result) {
    cache.putValue(email, { person: JSON.stringify(result.value) });
  }

  return result ? result.value : null;
}

/**
 * async promise, returns Graph contact and/or Person associated with the email provided
 * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
 *
 * @param {string} email
 * @returns {(Promise<Array<Person | Contact>>)}
 * @memberof Graph
 */
export function findUserByEmail(graph: IGraph, email: string): Promise<Array<Person | Contact>> {
  return Promise.all([findPeople(graph, email), findContactByEmail(graph, email)]).then(([people, contacts]) => {
    return ((people as any[]) || []).concat(contacts || []);
  });
}
