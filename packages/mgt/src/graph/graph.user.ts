/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { CacheItem, CacheSchema, CacheService } from '../utils/Cache';
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
 * Time to invalidate cache in ms
 * Currently set for 60 min
 */
const cacheInvalidationTime = 3600000;

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getMe(graph: IGraph): Promise<User> {
  const cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');
  const me = await cache.getValue('me');

  if (me && cacheInvalidationTime > Date.now() - me.timeCached) {
    return JSON.parse(me.user);
  }

  const graphRes = graph
    .api('me')
    .middlewareOptions(prepScopes('user.read'))
    .get();
  cache.putValue('me', { user: JSON.stringify(graphRes) });
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
  const cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');
  let person = null as IDynamicPerson;

  const user: CacheUser = await cache.getValue(userId || 'me');

  if (user && Date.now() - user.timeCached < cacheInvalidationTime) {
    return JSON.parse(user.user);
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
  cache.putValue(userId || 'me', { user: JSON.stringify(person) });

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
  const cache = CacheService.getCache<CacheUser>(cacheSchema, 'users');

  // check cache
  let user: CacheUser = await cache.getValue(userPrincipleName);
  // is it stored and is timestamp good?
  if (user || cacheInvalidationTime > Date.now() - user.timeCached) {
    // return without any worries
    return JSON.parse(user.user);
  }
  // else we must grab it
  user = await graph
    .api(`/users/${userPrincipleName}`)
    .middlewareOptions(prepScopes(scopes))
    .get();
  cache.putValue(userPrincipleName, { user: JSON.stringify(user) });
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

  for (const id of userIds) {
    if (id !== '') {
      batch.get(id, `/users/${id}`, ['user.readbasic.all']);
    }
  }

  try {
    const responses = await batch.executeAll();
    const people = [];

    // iterate over userIds to ensure the order of ids
    for (const id of userIds) {
      const response = responses.get(id);
      if (response && response.content) {
        people.push(response.content);
      }
    }

    return people;
  } catch (_) {
    // fallback to making the request one by one
    try {
      return Promise.all(userIds.filter(id => id && id !== '').map(id => getUser(graph, id)));
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
  const cache = CacheService.getCache<CacheUserQuery>(cacheSchema, 'usersQuery');
  if (!peopleQueries || peopleQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch();
  const people = [];
  let cacheRes: CacheUserQuery;

  for (const personQuery of peopleQueries) {
    cacheRes = await cache.getValue(personQuery);

    if (cacheRes && cacheInvalidationTime > Date.now() - cacheRes.timeCached) {
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
        cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(response.content.value[0])] });
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
              cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(personArray[0])] });
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
  const cache = CacheService.getCache<CacheUserQuery>(cacheSchema, 'usersQuery');
  const item = { maxResults: top, results: null };
  const result: CacheUserQuery = await cache.getValue(query);

  if (result && cacheInvalidationTime > Date.now() - result.timeCached) {
    return result.results.map(userStr => JSON.parse(userStr));
  }

  const graphResult = await graph
    .api('users')
    .filter(
      `startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`
    )
    .top(top)
    .middlewareOptions(prepScopes(scopes))
    .get();

  item.results = graphResult.value.map(userStr => JSON.stringify(userStr));
  if (item.results) {
    cache.putValue(query, item);
  }
  return item.results ? graphResult.value : null;
}
