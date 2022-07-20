/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { Group } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';

/**
 * Group Type enumeration
 *
 * @export
 * @enum {number}
 */
export enum GroupType {
  /**
   * Any group Type
   */
  any = 0,

  /**
   * Office 365 group
   */
  // tslint:disable-next-line:no-bitwise
  unified = 1 << 0,

  /**
   * Security group
   */
  // tslint:disable-next-line:no-bitwise
  security = 1 << 1,

  /**
   * Mail Enabled Security group
   */
  // tslint:disable-next-line:no-bitwise
  mailenabledsecurity = 1 << 2,

  /**
   * Distribution Group
   */
  // tslint:disable-next-line:no-bitwise
  distribution = 1 << 3
}

/**
 * Object to be stored in cache
 */
export interface CacheGroup extends CacheItem {
  /**
   * stringified json representing a user
   */
  group?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
interface CacheGroupQuery extends CacheItem {
  /**
   * json representing a person stored as string
   */
  groups?: string[];
  /**
   * top number of results
   */
  top?: number;
}

/**
 * Defines the expiration time
 */
const getGroupsInvalidationTime = (): number =>
  CacheService.config.groups.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the groups store is enabled
 */
const getIsGroupsCacheEnabled = (): boolean => CacheService.config.groups.isEnabled && CacheService.config.isEnabled;

/**
 * Searches the Graph for Groups
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {GroupType} [groupTypes=GroupType.any] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export async function findGroups(
  graph: IGraph,
  query: string,
  top: number = 10,
  groupTypes: GroupType = GroupType.any,
  groupFilters: string = ''
): Promise<Group[]> {
  const scopes = 'Group.Read.All';

  let cache: CacheStore<CacheGroupQuery>;
  const key = `${query ? query : '*'}*${groupTypes}*${groupFilters}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    const cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x)).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  let filterQuery = '';
  let responses;
  let batchedResult = [];

  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (groupFilters) {
    filterQuery += `${query ? ' and ' : ''}${groupFilters}`;
  }

  if (groupTypes !== GroupType.any) {
    const batch = graph.createBatch();

    const filterGroups = [];

    // tslint:disable-next-line:no-bitwise
    if (GroupType.unified === (groupTypes & GroupType.unified)) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.security === (groupTypes & GroupType.security)) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.mailenabledsecurity === (groupTypes & GroupType.mailenabledsecurity)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.distribution === (groupTypes & GroupType.distribution)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery = filterQuery ? `${filterQuery} and ` : '';
    for (let filter of filterGroups) {
      batch.get(filter, `/groups?$filter=${filterQuery + filter}`, ['Group.Read.All']);
    }

    try {
      responses = await batch.executeAll();

      for (let i = 0; i < filterGroups.length; i++) {
        if (responses.get(filterGroups[i]).content.value) {
          for (let group of responses.get(filterGroups[i]).content.value) {
            let repeat = batchedResult.filter(batchedGroup => batchedGroup.id === group.id);
            if (repeat.length === 0) {
              batchedResult.push(group);
            }
            repeat = [];
          }
        }
      }
    } catch (_) {
      try {
        let queries = [];
        for (let filter of filterGroups) {
          queries.push(
            await graph
              .api('groups')
              .filter(`${filterQuery} and ${filter}`)
              .top(top)
              .count(true)
              .header('ConsistencyLevel', 'eventual')
              .middlewareOptions(prepScopes(scopes))
              .get()
          );
        }
        return Promise.all(queries);
      } catch (_) {
        return [];
      }
    }
  } else {
    if (batchedResult.length === 0) {
      const result = await graph
        .api('groups')
        .filter(filterQuery)
        .top(top)
        .count(true)
        .header('ConsistencyLevel', 'eventual')
        .middlewareOptions(prepScopes(scopes))
        .get();
      if (getIsGroupsCacheEnabled() && result) {
        cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top: top });
      }
      return result ? result.value : null;
    }
  }

  return batchedResult;
}

/**
 * Searches the Graph for group members
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {string} groupId - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {boolean} [transitive=false] - whether the return should contain a flat list of all nested members
 * @param {GroupType} [groupTypes=GroupType.any] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export async function findGroupsFromGroup(
  graph: IGraph,
  query: string,
  groupId: string,
  top: number = 10,
  transitive: boolean = false,
  groupTypes: GroupType = GroupType.any
): Promise<Group[]> {
  const scopes = 'Group.Read.All';

  let cache: CacheStore<CacheGroupQuery>;
  const key = `${groupId}:${query || '*'}:${groupTypes}:${transitive}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    const cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x)).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  const apiUrl = `groups/${groupId}/${transitive ? 'transitiveMembers' : 'members'}/microsoft.graph.group`;
  let filterQuery = '';
  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (groupTypes !== GroupType.any) {
    const filterGroups = [];

    // tslint:disable-next-line:no-bitwise
    if (GroupType.unified === (groupTypes & GroupType.unified)) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.security === (groupTypes & GroupType.security)) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.mailenabledsecurity === (groupTypes & GroupType.mailenabledsecurity)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.distribution === (groupTypes & GroupType.distribution)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery += (query !== '' ? ' and ' : '') + filterGroups.join(' or ');
  }

  const result = await graph
    .api(apiUrl)
    .filter(filterQuery)
    .count(true)
    .top(top)
    .header('ConsistencyLevel', 'eventual')
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (getIsGroupsCacheEnabled() && result) {
    cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top: top });
  }

  return result ? result.value : null;
}

/**
 * async promise, returns all Graph groups associated with the id provided
 *
 * @param {string} id
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getGroup(graph: IGraph, id: string, requestedProps?: string[]): Promise<Group> {
  const scopes = 'Group.Read.All';
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
    // check cache
    const group = await cache.getValue(id);

    // is it stored and is timestamp good?
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      const cachedData = group.group ? JSON.parse(group.group) : null;
      const uniqueProps =
        requestedProps && cachedData ? requestedProps.filter(prop => !Object.keys(cachedData).includes(prop)) : null;

      // return without any worries
      if (!uniqueProps || uniqueProps.length <= 1) {
        return cachedData;
      }
    }
  }

  let apiString = `/groups/${id}`;
  if (requestedProps) {
    apiString = apiString + '?$select=' + requestedProps.toString();
  }

  // else we must grab it
  const response = await graph.api(apiString).middlewareOptions(prepScopes(scopes)).get();
  if (getIsGroupsCacheEnabled()) {
    cache.putValue(id, { group: JSON.stringify(response) });
  }
  return response;
}

/**
 * Returns a Promise of Graph Groups array associated with the groupIds array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} groupIds, an array of string ids
 * @returns {Promise<Group[]>}
 */
export async function getGroupsForGroupIds(graph: IGraph, groupIds: string[], filters: string = ''): Promise<Group[]> {
  if (!groupIds || groupIds.length === 0) {
    return [];
  }
  const batch = graph.createBatch();
  const groupDict = {};
  const notInCache = [];
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
  }

  for (const id of groupIds) {
    groupDict[id] = null;
    let group = null;
    if (getIsGroupsCacheEnabled()) {
      group = await cache.getValue(id);
    }
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      groupDict[id] = group.group ? JSON.parse(group.group) : null;
    } else if (id !== '') {
      let apiUrl: string = `/groups/${id}`;
      if (filters) {
        apiUrl = `${apiUrl}?$filters=${filters}`;
      }
      batch.get(id, apiUrl, ['Group.Read.All']);
      notInCache.push(id);
    }
  }
  try {
    const responses = await batch.executeAll();
    // iterate over groupIds to ensure the order of ids
    for (const id of groupIds) {
      const response = responses.get(id);
      if (response && response.content) {
        groupDict[id] = response.content;
        if (getIsGroupsCacheEnabled()) {
          cache.putValue(id, { group: JSON.stringify(response.content) });
        }
      }
    }
    return Promise.all(Object.values(groupDict));
  } catch (_) {
    // fallback to making the request one by one
    try {
      // call getGroup for all the users that weren't cached
      groupIds.filter(id => notInCache.includes(id)).forEach(id => (groupDict[id] = getGroup(graph, id)));
      if (getIsGroupsCacheEnabled()) {
        // store all users that weren't retrieved from the cache, into the cache
        groupIds
          .filter(id => notInCache.includes(id))
          .forEach(async id => cache.putValue(id, { group: JSON.stringify(await groupDict[id]) }));
      }
      return Promise.all(Object.values(groupDict));
    } catch (_) {
      return [];
    }
  }
}

/**
 * Gets groups from the graph that are in the group ids
 * @param graph
 * @param query
 * @param groupId
 * @param top
 * @param transitive
 * @param groupTypes
 * @param filters
 * @returns
 */
export async function findGroupsFromGroupIds(
  graph: IGraph,
  query: string,
  groupIds: string[],
  top: number = 10,
  groupTypes: GroupType = GroupType.any,
  filters: string = ''
): Promise<Group[]> {
  const foundGroups: Group[] = [];
  const graphGroups = await findGroups(graph, query, top, groupTypes, filters);
  if (graphGroups) {
    for (let i = 0; i < graphGroups.length; i++) {
      const group = graphGroups[i];
      if (group.id && groupIds.includes(group.id)) {
        foundGroups.push(group);
      }
    }
  }
  return foundGroups;
}
