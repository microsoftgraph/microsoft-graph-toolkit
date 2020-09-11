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
import {
  getPhotoForResource,
  getPhotoFromCache,
  getPhotoInvalidationTime,
  photosCacheEnabled,
  storePhotoInCache,
  CachePhoto
} from './graph.photos';
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
 * Name of the users store name
 */
const userStore: string = 'users';

/**
 * Name of the users query store name
 */
const queryStore: string = 'usersQuery';

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getMe(graph: IGraph): Promise<User> {
  let cache: CacheStore<CacheUser>;
  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, userStore);
    const me = await cache.getValue('me');

    if (me && getUserInvalidationTime() > Date.now() - me.timeCached) {
      return JSON.parse(me.user);
    }
  }
  const response = graph
    .api('me')
    .middlewareOptions(prepScopes('user.read'))
    .get();
  if (usersCacheEnabled()) {
    cache.putValue('me', { user: JSON.stringify(await response) });
  }
  return response;
}

/**
 * asnyc promise, returns IDynamicPerson
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */
export async function getUserWithPhoto(graph: IGraph, userId?: string): Promise<IDynamicPerson> {
  let person = null as IDynamicPerson;
  let cache: CacheStore<CacheUser>;
  let photo = null;
  let user: IDynamicPerson;
  let cachedPhoto: CachePhoto;
  const resource = userId ? `users/${userId}` : 'me';
  const scopes = userId ? ['user.readbasic.all'] : ['user.read'];

  // attempt to get user and photo from cache if enabled
  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, userStore);
    const cachedUser = await cache.getValue(userId || 'me');
    if (cachedUser && getUserInvalidationTime() > Date.now() - cachedUser.timeCached) {
      user = JSON.parse(cachedUser.user);
    }
  }
  if (photosCacheEnabled()) {
    cachedPhoto = await getPhotoFromCache(userId || 'me', 'users');
    if (cachedPhoto && getPhotoInvalidationTime() > Date.now() - cachedPhoto.timeCached) {
      photo = cachedPhoto.photo;
    } else if (cachedPhoto) {
      try {
        const response = await graph.api(`${resource}/photo`).get();
        if (response && response['@odata.mediaEtag'] && response['@odata.mediaEtag'] === cachedPhoto.eTag) {
          // put current image into the cache to update the timestamp since etag is the same
          storePhotoInCache(userId || 'me', 'users', cachedPhoto);
          photo = cachedPhoto.photo;
        }
      } catch (e) {}
    }
  }

  if (!photo && !user) {
    let eTag: string;

    // batch calls
    const batch = graph.createBatch();
    if (userId) {
      batch.get('user', `/users/${userId}`, ['user.readbasic.all']);
      batch.get('photo', `users/${userId}/photo/$value`, ['user.readbasic.all']);
    } else {
      batch.get('user', 'me', ['user.read']);
      batch.get('photo', 'me/photo/$value', ['user.read']);
    }
    const response = await batch.executeAll();

    const photoResponse = response.get('photo');
    if (photoResponse) {
      eTag = photoResponse.headers['ETag'];
      photo = photoResponse.content;
    }

    const userResponse = response.get('user');
    if (userResponse) {
      user = userResponse.content;
    }

    // store user & photo in their respective cache
    if (usersCacheEnabled()) {
      cache.putValue(userId || 'me', { user: JSON.stringify(user) });
    }
    if (photosCacheEnabled()) {
      storePhotoInCache(userId || 'me', 'users', { eTag, photo: photo });
    }
  } else if (!photo) {
    // get photo from graph
    const resource = userId ? `users/${userId}` : 'me';
    const scopes = userId ? ['user.readbasic.all'] : ['user.read'];
    const response = await getPhotoForResource(graph, resource, scopes);
    if (response) {
      if (photosCacheEnabled()) {
        storePhotoInCache(userId || 'me', 'users', { eTag: response.eTag, photo: response.photo });
      }
      photo = response.photo;
    }
  } else if (!user) {
    // get user from graph
    const response = userId
      ? await graph
          .api(`/users/${userId}`)
          .middlewareOptions(prepScopes('user.readbasic.all'))
          .get()
      : await graph
          .api('me')
          .middlewareOptions(prepScopes('user.read'))
          .get();
    if (response) {
      if (usersCacheEnabled()) {
        cache.putValue(userId || 'me', { user: JSON.stringify(response) });
      }
      user = response;
    }
  }
  if (user) {
    person = user;
    person.personImage = photo;
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

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, userStore);
    // check cache
    const user = await cache.getValue(userPrincipleName);
    // is it stored and is timestamp good?
    if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
      // return without any worries
      return JSON.parse(user.user);
    }
  }
  // else we must grab it
  const response = await graph
    .api(`/users/${userPrincipleName}`)
    .middlewareOptions(prepScopes(scopes))
    .get();
  if (usersCacheEnabled()) {
    cache.putValue(userPrincipleName, { user: JSON.stringify(response) });
  }
  return response;
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
  const peopleDict = {};
  const notInCache = [];
  let cache: CacheStore<CacheUser>;

  if (usersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(cacheSchema, userStore);
  }

  for (const id of userIds) {
    peopleDict[id] = null;
    let user = null;
    if (usersCacheEnabled()) {
      user = await cache.getValue(id);
    }
    if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
      peopleDict[id] = JSON.parse(user.user);
    } else if (id !== '') {
      batch.get(id, `/users/${id}`, ['user.readbasic.all']);
      notInCache.push(id);
    }
  }
  try {
    const responses = await batch.executeAll();
    // iterate over userIds to ensure the order of ids
    for (const id of userIds) {
      const response = responses.get(id);
      if (response && response.content) {
        peopleDict[id] = response.content;
        if (usersCacheEnabled()) {
          cache.putValue(id, { user: JSON.stringify(response.content) });
        }
      }
    }
    return Promise.all(Object.values(peopleDict));
  } catch (_) {
    // fallback to making the request one by one
    try {
      // call getUser for all the users that weren't cached
      userIds.filter(id => notInCache.includes(id)).forEach(id => (peopleDict[id] = getUser(graph, id)));
      if (usersCacheEnabled()) {
        // store all users that weren't retrieved from the cache, into the cache
        userIds
          .filter(id => notInCache.includes(id))
          .forEach(async id => cache.putValue(id, { user: JSON.stringify(await peopleDict[id]) }));
      }
      return Promise.all(Object.values(peopleDict));
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
    cache = CacheService.getCache<CacheUserQuery>(cacheSchema, queryStore);
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
    cache = CacheService.getCache<CacheUserQuery>(cacheSchema, queryStore);
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
