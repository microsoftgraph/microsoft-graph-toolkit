/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { User } from '@microsoft/microsoft-graph-types';
import { CacheItem, CacheSchema, CacheService, CacheStore } from '../utils/Cache';
import { prepScopes } from '../utils/GraphHelpers';
import { findPeople } from './graph.people';
import { IDynamicPerson } from './types';

/**
 * Describes the organization of the cache db
 */
const cacheSchema: CacheSchema = {
  name: 'users',
  stores: {
    users: {},
    usersQuery: {}
  },
  version: 1
};

/**
 * Object to be stored in cache
 */
interface CacheUser extends CacheItem {
  /**
   * stringified json representing a user
   */
  user?: string;
}

/**
 * Object to be stored in cache
 */
interface CacheUserQuery extends CacheItem {
  /**
   * max number of results the query asks for
   */
  maxResults?: number;
  /**
   * list of users returned by query
   */
  results?: string[];
}

/**
 * Defines the time it takes for objects in the cache to expire
 */
const getUserInvalidationTime = (): number =>
  CacheService.config.users.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
const usersCacheEnabled = (): boolean => CacheService.config.users.isEnabled && CacheService.config.isEnabled;

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getMe(graph: IGraph): Promise<User> {
  let cache: CacheStore<CacheUser>;
  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');
    const me = await cache.getValue('me');

    if (me && getUserInvalidationTime() > Date.now() - me.timeCached) {
      return JSON.parse(me.user);
    }
  }
  const graphRes = graph
    .api('me')
    .middlewareOptions(prepScopes('user.read'))
    .get();
  if (usersCacheEnabled()) {
    cache.putValue('me', { user: JSON.stringify(graphRes) });
  }
  return graphRes;
}

/**
 * asnyc promise, returns IDynamicPerson
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */

export async function getUserWithPhoto(graph: IGraph, userId?: string): Promise<IDynamicPerson> {
  const batch = graph.createBatch();
  let person = null as IDynamicPerson;
  let cache: CacheStore<CacheUser>;

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');
    const user: CacheUser = await cache.getValue(userId || 'me');

    if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
      return JSON.parse(user.user);
    }
  }

  if (userId) {
    batch.get('user', `/users/${userId}`, ['user.readbasic.all']);
    batch.get('photo', `users/${userId}/photo/$value`, ['user.readbasic.all']);
  } else {
    batch.get('user', 'me', ['user.read']);
    batch.get('photo', 'me/photo/$value', ['user.read']);
  }
  const response = await batch.executeAll();

  const photoResponse = response.get('photo');
  const userDetailsResponse = response.get('user');

  if (userDetailsResponse.content) {
    person = userDetailsResponse.content;

    person.personImage = photoResponse && photoResponse.content;
  }

  if (usersCacheEnabled()) {
    cache.putValue(userId || 'me', { user: JSON.stringify(person) });
  }

  return person;
}

/**
 * async promise, returns all Graph users associated with the userPrincipleName provided
 *
 * @param {string} userPrincipleName
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getUser(graph: IGraph, userPrincipleName: string): Promise<User> {
  const scopes = 'user.readbasic.all';
  let cache: CacheStore<CacheUser>;
  let user: CacheUser;

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');
    // check cache
    user = await cache.getValue(userPrincipleName);
    // is it stored and is timestamp good?
    if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
      // return without any worries
      return JSON.parse(user.user);
    }
  }
  // else we must grab it
  user = await graph
    .api(`/users/${userPrincipleName}`)
    .middlewareOptions(prepScopes(scopes))
    .get();
  if (usersCacheEnabled()) {
    cache.putValue(userPrincipleName, { user: JSON.stringify(user) });
  }
  return JSON.parse(user.user);
}

/**
 * Returns a Promise of Graph Users array associated with the user ids array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} userIds, an array of string ids
 * @returns {Promise<User[]>}
 */
export async function getUsersForUserIds(graph: IGraph, userIds: string[]): Promise<User[]> {
  if (!userIds || userIds.length === 0) {
    return [];
  }
  const batch = graph.createBatch();
  let people = [];
  let cache: CacheStore<CacheUser>;

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, 'usersQuery');
    for (const id of userIds) {
      const user = await cache.getValue(id);
      if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
        people.push(JSON.parse(user.user));
      } else if (id !== '') {
        batch.get(id, `/users/${id}`, ['user.readbasic.all']);
      }
    }
  }
  try {
    const responses = await batch.executeAll();

    // iterate over userIds to ensure the order of ids
    for (const id of userIds) {
      const response = responses.get(id);
      if (response && response.content) {
        people.push(response.content);
        if (usersCacheEnabled()) {
          cache.putValue(id, { user: JSON.stringify(response.content) });
        }
      }
    }

    return people;
  } catch (_) {
    // fallback to making the request one by one
    try {
      people = userIds.filter(id => id && id !== '').map(id => getUser(graph, id));
      if (usersCacheEnabled()) {
        people.forEach((u: User) => cache.putValue(u.id, { user: JSON.stringify(u) }));
      }
      return Promise.all(people);
    } catch (_) {
      return [];
    }
  }
}

/**
 * Returns a Promise of Graph Users array associated with the people queries array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} peopleQueries, an array of string ids
 * @returns {Promise<User[]>}
 */
export async function getUsersForPeopleQueries(graph: IGraph, peopleQueries: string[]): Promise<User[]> {
  if (!peopleQueries || peopleQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch();
  const people = [];
  let cacheRes: CacheUserQuery;
  let cache: CacheStore<CacheUserQuery>;
  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(cacheSchema, 'usersQuery');
  }

  for (const personQuery of peopleQueries) {
    if (usersCacheEnabled()) {
      cacheRes = await cache.getValue(personQuery);
    }

    if (usersCacheEnabled() && cacheRes && getUserInvalidationTime() > Date.now() - cacheRes.timeCached) {
      people.push(JSON.parse(cacheRes.results[0]));
    } else if (personQuery !== '') {
      batch.get(personQuery, `/me/people?$search="${personQuery}"`, ['people.read']);
    }
  }

  try {
    const responses = await batch.executeAll();

    for (const personQuery of peopleQueries) {
      const response = responses.get(personQuery);
      if (response && response.content && response.content.value && response.content.value.length > 0) {
        people.push(response.content.value[0]);
        if (usersCacheEnabled()) {
          cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(response.content.value[0])] });
        }
      }
    }

    return people;
  } catch (_) {
    try {
      return Promise.all(
        peopleQueries
          .filter(personQuery => personQuery && personQuery !== '')
          .map(async personQuery => {
            const personArray = await findPeople(graph, personQuery, 1);
            if (personArray && personArray.length) {
              if (usersCacheEnabled()) {
                cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(personArray[0])] });
              }
              return personArray[0];
            }
          })
      );
    } catch (_) {
      return [];
    }
  }
}

/**
 * Search Microsoft Graph for Users in the organization
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - the string to search for
 * @param {number} [top=10] - maximum number of results to return
 * @returns {Promise<User[]>}
 */
export async function findUsers(graph: IGraph, query: string, top: number = 10): Promise<User[]> {
  const scopes = 'User.ReadBasic.All';
  const item = { maxResults: top, results: null };
  let cache: CacheStore<CacheUserQuery>;

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(cacheSchema, 'usersQuery');
    const result: CacheUserQuery = await cache.getValue(query);

    if (result && getUserInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(userStr => JSON.parse(userStr));
    }
  }

  const graphResult = await graph
    .api('users')
    .filter(
      `startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`
    )
    .top(top)
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (usersCacheEnabled() && graphResult) {
    item.results = graphResult.value.map(userStr => JSON.stringify(userStr));
    cache.putValue(query, item);
  }
  return graphResult ? graphResult.value : null;
}
