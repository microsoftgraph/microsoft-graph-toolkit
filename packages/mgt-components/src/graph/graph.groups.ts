/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  BatchResponse,
  CacheItem,
  CacheService,
  CacheStore,
  CollectionResponse,
  IGraph,
  prepScopes
} from '@microsoft/mgt-element';
import { Group } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';

const groupTypeValues = ['any', 'unified', 'security', 'mailenabledsecurity', 'distribution'] as const;

/**
 * Group Type enumeration
 *
 * @export
 * @enum {string}
 */
export type GroupType = (typeof groupTypeValues)[number];
export const isGroupType = (groupType: string): groupType is (typeof groupTypeValues)[number] =>
  groupTypeValues.includes(groupType as GroupType);
export const groupTypeConverter = (value: string, defaultValue: GroupType = 'any'): GroupType =>
  isGroupType(value) ? value : defaultValue;

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

const validGroupQueryScopes = [
  'GroupMember.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'Group.ReadWrite.All',
  'Directory.ReadWrite.All'
];

const validTransitiveGroupMemberScopes = [
  'GroupMember.Read.All',
  'Group.Read.All',
  'Directory.Read.All',
  'GroupMember.ReadWrite.All',
  'Group.ReadWrite.All'
];

/**
 * Searches the Graph for Groups
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {GroupType} [groupTypes=["any"]] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export const findGroups = async (
  graph: IGraph,
  query: string,
  top = 10,
  groupTypes: GroupType[] = ['any'],
  groupFilters = ''
): Promise<Group[]> => {
  const groupTypesString = Array.isArray(groupTypes) ? groupTypes.join('+') : JSON.stringify(groupTypes);
  let cache: CacheStore<CacheGroupQuery>;
  const key = `${query ? query : '*'}*${groupTypesString}*${groupFilters}:${top}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    const cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x) as Group).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  let filterQuery = '';
  let responses: Map<string, BatchResponse<CollectionResponse<Group>>>;
  const batchedResult: Group[] = [];

  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (groupFilters) {
    filterQuery += `${query ? ' and ' : ''}${groupFilters}`;
  }

  if (!groupTypes.includes('any')) {
    const batch = graph.createBatch<CollectionResponse<Group>>();

    const filterGroups: string[] = [];

    if (groupTypes.includes('unified')) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    if (groupTypes.includes('security')) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    if (groupTypes.includes('mailenabledsecurity')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    if (groupTypes.includes('distribution')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery = filterQuery ? `${filterQuery} and ` : '';
    for (const filter of filterGroups) {
      const fullUrl = `/groups?$filter=${filterQuery + filter}&$top=${top}`;
      batch.get(filter, fullUrl, validGroupQueryScopes);
    }

    try {
      responses = await batch.executeAll();

      for (const filterGroup of filterGroups) {
        if (responses.get(filterGroup).content.value) {
          for (const group of responses.get(filterGroup).content.value) {
            const repeat = batchedResult.find(batchedGroup => batchedGroup.id === group.id);
            if (!repeat) {
              batchedResult.push(group);
            }
          }
        }
      }
    } catch (_) {
      try {
        const queries: Promise<CollectionResponse<Group>>[] = [];
        for (const filter of filterGroups) {
          queries.push(
            graph
              .api('groups')
              .filter(`${filterQuery} and ${filter}`)
              .top(top)
              .count(true)
              .header('ConsistencyLevel', 'eventual')
              .middlewareOptions(prepScopes(validGroupQueryScopes))
              .get() as Promise<CollectionResponse<Group>>
          );
        }
        return (await Promise.all(queries)).map(x => x.value).reduce((a, b) => a.concat(b), []);
      } catch (e) {
        return [];
      }
    }
  } else {
    if (batchedResult.length === 0) {
      const result = (await graph
        .api('groups')
        .filter(filterQuery)
        .top(top)
        .count(true)
        .header('ConsistencyLevel', 'eventual')
        .middlewareOptions(prepScopes(validGroupQueryScopes))
        .get()) as CollectionResponse<Group>;
      if (getIsGroupsCacheEnabled() && result) {
        await cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top });
      }
      return result ? result.value : null;
    }
  }

  return batchedResult;
};

/**
 * Searches the Graph for group members
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {string} groupId - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {boolean} [transitive=false] - whether the return should contain a flat list of all nested members
 * @param {GroupType} [groupTypes=["any"]] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export const findGroupsFromGroup = async (
  graph: IGraph,
  query: string,
  groupId: string,
  top = 10,
  transitive = false,
  groupTypes: GroupType[] = ['any']
): Promise<Group[]> => {
  let cache: CacheStore<CacheGroupQuery>;
  const groupTypesString = Array.isArray(groupTypes) ? groupTypes.join('+') : JSON.stringify(groupTypes);
  const key = `${groupId}:${query || '*'}:${groupTypesString}:${transitive}`;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groupsQuery);
    const cacheGroupQuery = await cache.getValue(key);
    if (cacheGroupQuery && getGroupsInvalidationTime() > Date.now() - cacheGroupQuery.timeCached) {
      if (cacheGroupQuery.top >= top) {
        // if request is less than the cache's requests, return a slice of the results
        return cacheGroupQuery.groups.map(x => JSON.parse(x) as Group).slice(0, top + 1);
      }
      // if the new request needs more results than what's presently in the cache, graph must be called again
    }
  }

  const apiUrl = `groups/${groupId}/${transitive ? 'transitiveMembers' : 'members'}/microsoft.graph.group`;
  let filterQuery = '';
  if (query !== '') {
    filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;
  }

  if (!groupTypes.includes('any')) {
    const filterGroups = [];

    if (groupTypes.includes('unified')) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    if (groupTypes.includes('security')) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    if (groupTypes.includes('mailenabledsecurity')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    if (groupTypes.includes('distribution')) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery += (query !== '' ? ' and ' : '') + filterGroups.join(' or ');
  }

  const result = (await graph
    .api(apiUrl)
    .filter(filterQuery)
    .count(true)
    .top(top)
    .header('ConsistencyLevel', 'eventual')
    .middlewareOptions(prepScopes(validTransitiveGroupMemberScopes))
    .get()) as CollectionResponse<Group>;

  if (getIsGroupsCacheEnabled() && result) {
    await cache.putValue(key, { groups: result.value.map(x => JSON.stringify(x)), top });
  }

  return result ? result.value : null;
};

/**
 * async promise, returns all Graph groups associated with the id provided
 *
 * @param {string} id
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export const getGroup = async (graph: IGraph, id: string, requestedProps?: string[]): Promise<Group> => {
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
    // check cache
    const group = await cache.getValue(id);

    // is it stored and is timestamp good?
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      const cachedData = group.group ? (JSON.parse(group.group) as Group) : null;
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
  const response = (await graph.api(apiString).middlewareOptions(prepScopes(validGroupQueryScopes)).get()) as Group;
  if (getIsGroupsCacheEnabled()) {
    await cache.putValue(id, { group: JSON.stringify(response) });
  }
  return response;
};

/**
 * Returns a Promise of Graph Groups array associated with the groupIds array
 *
 * @export
 * @param {IGraph} graph
 * @param {string[]} groupIds, an array of string ids
 * @returns {Promise<Group[]>}
 */
export const getGroupsForGroupIds = async (graph: IGraph, groupIds: string[], filters = ''): Promise<Group[]> => {
  if (!groupIds || groupIds.length === 0) {
    return [];
  }
  const batch = graph.createBatch();
  const groupDict: Record<string, Group | Promise<Group>> = {};
  const notInCache: string[] = [];
  let cache: CacheStore<CacheGroup>;

  if (getIsGroupsCacheEnabled()) {
    cache = CacheService.getCache(schemas.groups, schemas.groups.stores.groups);
  }

  for (const id of groupIds) {
    groupDict[id] = null;
    let group: CacheGroup;
    if (getIsGroupsCacheEnabled()) {
      group = await cache.getValue(id);
    }
    if (group && getGroupsInvalidationTime() > Date.now() - group.timeCached) {
      groupDict[id] = group.group ? (JSON.parse(group.group) as Group) : null;
    } else if (id !== '') {
      let apiUrl = `/groups/${id}`;
      if (filters) {
        apiUrl = `${apiUrl}?$filters=${filters}`;
      }
      batch.get(id, apiUrl, validGroupQueryScopes);
      notInCache.push(id);
    }
  }
  try {
    const responses = await batch.executeAll();
    // iterate over groupIds to ensure the order of ids
    for (const id of groupIds) {
      const response = responses.get(id);
      if (response?.content) {
        groupDict[id] = response.content as Group;
        if (getIsGroupsCacheEnabled()) {
          await cache.putValue(id, { group: JSON.stringify(response.content) });
        }
      }
    }
    return Promise.all(Object.values(groupDict));
  } catch (_) {
    // fallback to making the request one by one
    try {
      // call getGroup for all the users that weren't cached
      groupIds
        .filter(id => notInCache.includes(id))
        .forEach(id => {
          groupDict[id] = getGroup(graph, id);
        });
      if (getIsGroupsCacheEnabled()) {
        // store all users that weren't retrieved from the cache, into the cache
        await Promise.all(
          groupIds
            .filter(id => notInCache.includes(id))
            .map(async id => await cache.putValue(id, { group: JSON.stringify(await groupDict[id]) }))
        );
      }
      return Promise.all(Object.values(groupDict));
    } catch (e) {
      return [];
    }
  }
};

/**
 * Gets groups from the graph that are in the group ids
 *
 * @param graph
 * @param query
 * @param groupId
 * @param top
 * @param transitive
 * @param groupTypes
 * @param filters
 * @returns
 */
export const findGroupsFromGroupIds = async (
  graph: IGraph,
  query: string,
  groupIds: string[],
  top = 10,
  groupTypes: GroupType[] = ['any'],
  filters = ''
): Promise<Group[]> => {
  const foundGroups: Group[] = [];
  const graphGroups = await findGroups(graph, query, top, groupTypes, filters);
  if (graphGroups) {
    for (const group of graphGroups) {
      if (group.id && groupIds.includes(group.id)) {
        foundGroups.push(group);
      }
    }
  }
  return foundGroups;
};
