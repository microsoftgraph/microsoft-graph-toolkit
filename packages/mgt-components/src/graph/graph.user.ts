/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheService, CacheStore, CollectionResponse, IGraph, prepScopes } from '@microsoft/mgt-element';
import { User } from '@microsoft/microsoft-graph-types';

import { GraphRequest } from '@microsoft/microsoft-graph-client';
import { schemas } from './cacheStores';
import { findPeople, PersonType } from './graph.people';
import { IDynamicPerson } from './types';

/**
 * Object to be stored in cache
 */
export interface CacheUser extends CacheItem {
  /**
   * stringified json representing a user
   */
  user?: string;
}

/**
 * Object to be stored in cache
 */
export interface CacheUserQuery extends CacheItem {
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
export const getUserInvalidationTime = (): number =>
  CacheService.config.users.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
export const getIsUsersCacheEnabled = (): boolean =>
  CacheService.config.users.isEnabled && CacheService.config.isEnabled;

export const getUsers = async (graph: IGraph, userFilters = '', top = 10): Promise<User[]> => {
  const apiString = '/users';
  let cache: CacheStore<CacheUserQuery>;
  const cacheKey = `${userFilters === '' ? '*' : userFilters}:${top}`;
  const cacheItem = { maxResults: top, results: null };

  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(schemas.users, schemas.users.stores.userFilters);
    const cacheRes = await cache.getValue(cacheKey);
    if (cacheRes && getUserInvalidationTime() > Date.now() - cacheRes.timeCached) {
      return cacheRes.results.map(userStr => JSON.parse(userStr) as User);
    }
  }
  const graphClient: GraphRequest = graph.api(apiString).top(top);

  if (userFilters) {
    graphClient.filter(userFilters);
  }

  try {
    const response = (await graphClient.middlewareOptions(prepScopes('user.read')).get()) as CollectionResponse<User>;
    if (getIsUsersCacheEnabled() && response) {
      cacheItem.results = response.value.map(userStr => JSON.stringify(userStr));
      await cache.putValue(userFilters, cacheItem);
    }
    return response.value;
    // eslint-disable-next-line no-empty
  } catch (error) {}
};

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export const getMe = async (graph: IGraph, requestedProps?: string[]): Promise<User> => {
  let cache: CacheStore<CacheUser>;
  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
    const me = await cache.getValue('me');

    if (me && getUserInvalidationTime() > Date.now() - me.timeCached) {
      const cachedData = JSON.parse(me.user) as User;
      const uniqueProps = requestedProps
        ? requestedProps.filter(prop => !Object.keys(cachedData).includes(prop))
        : null;

      // if requestedProps doesn't contain any unique props other than "@odata.context"
      if (!uniqueProps || uniqueProps.length <= 1) {
        return cachedData;
      }
    }
  }

  let apiString = 'me';
  if (requestedProps) {
    apiString = apiString + '?$select=' + requestedProps.toString();
  }
  const response = (await graph.api(apiString).middlewareOptions(prepScopes('user.read')).get()) as User;
  if (getIsUsersCacheEnabled()) {
    await cache.putValue('me', { user: JSON.stringify(response) });
  }
  return response;
};

/**
 * async promise, returns all Graph users associated with the userPrincipleName provided
 *
 * @param {string} userPrincipleName
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export const getUser = async (graph: IGraph, userPrincipleName: string, requestedProps?: string[]): Promise<User> => {
  const scopes = 'user.readbasic.all';
  let cache: CacheStore<CacheUser>;

  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
    // check cache
    const user = await cache.getValue(userPrincipleName);

    // is it stored and is timestamp good?
    if (user && getUserInvalidationTime() > Date.now() - user.timeCached) {
      const cachedData = user.user ? (JSON.parse(user.user) as User) : null;
      const uniqueProps =
        requestedProps && cachedData ? requestedProps.filter(prop => !Object.keys(cachedData).includes(prop)) : null;

      // return without any worries
      if (!uniqueProps || uniqueProps.length <= 1) {
        return cachedData;
      }
    }
  }

  let apiString = `/users/${userPrincipleName}`;
  if (requestedProps) {
    apiString = apiString + '?$select=' + requestedProps.toString();
  }

  // else we must grab it
  let response: User;
  try {
    response = (await graph.api(apiString).middlewareOptions(prepScopes(scopes)).get()) as User;
    // eslint-disable-next-line no-empty
  } catch (_) {}

  if (getIsUsersCacheEnabled()) {
    await cache.putValue(userPrincipleName, { user: JSON.stringify(response) });
  }
  return response;
};

/**
 * Returns a Promise of Graph Users array associated with the user ids array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} userIds, an array of string ids
 * @returns {Promise<User[]>}
 */
export const getUsersForUserIds = async (
  graph: IGraph,
  userIds: string[],
  searchInput = '',
  userFilters = '',
  fallbackDetails?: IDynamicPerson[]
): Promise<User[]> => {
  if (!userIds || userIds.length === 0) {
    return [];
  }
  const batch = graph.createBatch<User>();
  const peopleDict: Record<string, User | Promise<User>> = {};
  const peopleSearchMatches = {};
  const notInCache = [];
  searchInput = searchInput.toLowerCase();
  let cache: CacheStore<CacheUser>;

  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
  }

  for (const id of userIds) {
    peopleDict[id] = null;
    let apiUrl = `/users/${id}`;
    let user: User;
    let cacheUser: CacheUser;
    if (getIsUsersCacheEnabled()) {
      cacheUser = await cache.getValue(id);
    }
    if (cacheUser?.user && getUserInvalidationTime() > Date.now() - cacheUser.timeCached) {
      user = JSON.parse(cacheUser?.user) as User;

      if (searchInput) {
        if (user) {
          const displayName = user.displayName;
          const searchMatches = displayName?.toLowerCase().includes(searchInput);
          if (searchMatches) {
            peopleSearchMatches[id] = user;
          }
        }
      } else {
        if (user) {
          peopleDict[id] = user;
        } else {
          batch.get(id, apiUrl, ['user.readbasic.all']);
          notInCache.push(id);
        }
      }
    } else if (id !== '') {
      if (id.toString() === 'me') {
        peopleDict[id] = await getMe(graph);
      } else {
        apiUrl = `/users/${id}`;
        if (userFilters) {
          apiUrl += `${apiUrl}?$filter=${userFilters}`;
        }
        batch.get(id, apiUrl, ['user.readbasic.all']);
        notInCache.push(id);
      }
    }
  }
  try {
    if (batch.hasRequests) {
      const responses = await batch.executeAll();
      // iterate over userIds to ensure the order of ids
      for (const id of userIds) {
        const response = responses.get(id);
        if (response?.content) {
          const user = response.content;
          if (searchInput) {
            const displayName = user?.displayName.toLowerCase() || '';
            if (displayName.includes(searchInput)) {
              peopleSearchMatches[id] = user;
            }
          } else {
            peopleDict[id] = user;
          }

          if (getIsUsersCacheEnabled()) {
            await cache.putValue(id, { user: JSON.stringify(user) });
          }
        } else {
          const fallback = fallbackDetails.find(detail => Object.values(detail).includes(id));
          if (fallback) {
            peopleDict[id] = fallback as User;
          }
        }
      }
    }
    if (searchInput && Object.keys(peopleSearchMatches).length) {
      return Promise.all(Object.values(peopleSearchMatches));
    }
    return Promise.all(Object.values(peopleDict));
  } catch (_) {
    // fallback to making the request one by one
    try {
      // call getUser for all the users that weren't cached
      userIds
        .filter(id => notInCache.includes(id))
        .forEach(id => {
          peopleDict[id] = getUser(graph, id);
        });
      if (getIsUsersCacheEnabled()) {
        // store all users that weren't retrieved from the cache, into the cache
        await Promise.all(
          userIds
            .filter(id => notInCache.includes(id))
            .map(async id => await cache.putValue(id, { user: JSON.stringify(await peopleDict[id]) }))
        );
      }
      return Promise.all(Object.values(peopleDict));
    } catch (e) {
      return [];
    }
  }
};

/**
 * Returns a Promise of Graph Users array associated with the people queries array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} peopleQueries, an array of string ids
 * @returns {Promise<User[]>}
 */
export const getUsersForPeopleQueries = async (
  graph: IGraph,
  peopleQueries: string[],
  fallbackDetails?: IDynamicPerson[]
): Promise<User[]> => {
  if (!peopleQueries || peopleQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch<CollectionResponse<User>>();
  const people: User[] = [];
  let cacheRes: CacheUserQuery;
  let cache: CacheStore<CacheUserQuery>;
  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(schemas.users, schemas.users.stores.usersQuery);
  }

  for (const personQuery of peopleQueries) {
    if (getIsUsersCacheEnabled()) {
      cacheRes = await cache.getValue(personQuery);
    }

    if (
      getIsUsersCacheEnabled() &&
      cacheRes?.results[0] &&
      getUserInvalidationTime() > Date.now() - cacheRes.timeCached
    ) {
      const person = JSON.parse(cacheRes.results[0]) as User;
      people.push(person);
    } else {
      batch.get(personQuery, `/me/people?$search="${personQuery}"`, ['people.read'], {
        'X-PeopleQuery-QuerySources': 'Mailbox,Directory'
      });
    }
  }

  if (batch.hasRequests) {
    try {
      const responses = await batch.executeAll();

      for (const personQuery of peopleQueries) {
        const response = responses.get(personQuery);
        if (response?.content?.value && response.content.value.length > 0) {
          people.push(response.content.value[0]);
          if (getIsUsersCacheEnabled()) {
            await cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(response.content.value[0])] });
          }
        } else {
          const fallback = fallbackDetails.find(detail => Object.values(detail).includes(personQuery));
          if (fallback) {
            people.push(fallback as User);
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
              if (personArray?.length) {
                if (getIsUsersCacheEnabled()) {
                  await cache.putValue(personQuery, { maxResults: 1, results: [JSON.stringify(personArray[0])] });
                }
                return personArray[0];
              }
            })
        );
      } catch (e) {
        return [];
      }
    }
  }
  return people;
};

/**
 * Search Microsoft Graph for Users in the organization
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - the string to search for
 * @param {number} [top=10] - maximum number of results to return
 * @returns {Promise<User[]>}
 */
export const findUsers = async (graph: IGraph, query: string, top = 10, userFilters = ''): Promise<User[]> => {
  const scopes = 'User.ReadBasic.All';
  const item = { maxResults: top, results: null };
  const cacheKey = `${query}:${top}:${userFilters}`;
  let cache: CacheStore<CacheUserQuery>;

  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(schemas.users, schemas.users.stores.usersQuery);
    const result: CacheUserQuery = await cache.getValue(cacheKey);

    if (result && getUserInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(userStr => JSON.parse(userStr) as User);
    }
  }

  const encodedQuery = `${query.replace(/#/g, '%2523')}`;
  const graphBuilder = graph
    .api('users')
    .header('ConsistencyLevel', 'eventual')
    .count(true)
    .search(`"displayName:${encodedQuery}" OR "mail:${encodedQuery}"`);
  let graphResult: CollectionResponse<User>;

  if (userFilters !== '') {
    graphBuilder.filter(userFilters);
  }
  try {
    graphResult = (await graphBuilder.top(top).middlewareOptions(prepScopes(scopes)).get()) as CollectionResponse<User>;
    // eslint-disable-next-line no-empty
  } catch {}

  if (getIsUsersCacheEnabled() && graphResult) {
    item.results = graphResult.value.map(userStr => JSON.stringify(userStr));
    await cache.putValue(query, item);
  }
  return graphResult ? graphResult.value : null;
};

/**
 * async promise, returns all matching Graph users who are member of the specified group
 *
 * @param {string} query
 * @param {string} groupId - the group to query
 * @param {number} [top=10] - number of people to return
 * @param {PersonType} [personType=PersonType.person] - the type of person to search for
 * @param {boolean} [transitive=false] - whether the return should contain a flat list of all nested members
 * @returns {(Promise<User[]>)}
 */
export const findGroupMembers = async (
  graph: IGraph,
  query: string,
  groupId: string,
  top = 10,
  personType: PersonType = PersonType.person,
  transitive = false,
  userFilters = '',
  peopleFilters = ''
): Promise<User[]> => {
  const scopes = ['GroupMember.Read.All'];
  const item = { maxResults: top, results: null };

  let cache: CacheStore<CacheUserQuery>;
  const key = `${groupId || '*'}:${query || '*'}:${top}:${personType}:${transitive}:${userFilters}`;

  if (getIsUsersCacheEnabled()) {
    cache = CacheService.getCache<CacheUserQuery>(schemas.users, schemas.users.stores.usersQuery);
    const result: CacheUserQuery = await cache.getValue(key);

    if (result && getUserInvalidationTime() > Date.now() - result.timeCached) {
      return result.results.map(userStr => JSON.parse(userStr) as User);
    }
  }

  let filter = '';
  if (query) {
    filter = `startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}') or startswith(mail,'${query}') or startswith(userPrincipalName,'${query}')`;
  }

  let apiUrl = `/groups/${groupId}/${transitive ? 'transitiveMembers' : 'members'}`;
  if (personType === PersonType.person) {
    apiUrl += '/microsoft.graph.user';
  } else if (personType === PersonType.group) {
    apiUrl += '/microsoft.graph.group';
    if (query) {
      filter = `startswith(displayName,'${query}') or startswith(mail,'${query}')`;
    }
  }

  if (userFilters) {
    filter += query ? ` and ${userFilters}` : userFilters;
  }

  if (peopleFilters) {
    filter += query ? ` and ${peopleFilters}` : peopleFilters;
  }

  const graphResult = (await graph
    .api(apiUrl)
    .count(true)
    .top(top)
    .filter(filter)
    .header('ConsistencyLevel', 'eventual')
    .middlewareOptions(prepScopes(...scopes))
    .get()) as CollectionResponse<User>;

  if (getIsUsersCacheEnabled() && graphResult) {
    item.results = graphResult.value.map(userStr => JSON.stringify(userStr));
    await cache.putValue(key, item);
  }

  return graphResult ? graphResult.value : null;
};

export const findUsersFromGroupIds = async (
  graph: IGraph,
  query: string,
  groupIds: string[],
  top = 10,
  personType: PersonType = PersonType.person,
  transitive = false,
  groupFilters = ''
): Promise<User[]> => {
  const users: User[] = [];
  for (const groupId of groupIds) {
    try {
      const groupUsers = await findGroupMembers(graph, query, groupId, top, personType, transitive, groupFilters);
      users.push(...groupUsers);
    } catch (_) {
      continue;
    }
  }
  return users;
};
