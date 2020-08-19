/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Presence } from '@microsoft/microsoft-graph-types-beta';
import { BetaGraph } from '../BetaGraph';
import { CacheItem, CacheSchema, CacheService, CacheStore } from '../utils/Cache';
import { prepScopes } from '../utils/GraphHelpers';
import { IDynamicPerson } from './types';

/**
 * Definition of cache structure
 */
const cacheSchema: CacheSchema = {
  name: 'presence',
  stores: {
    presence: {}
  },
  version: 1
};

/**
 * Object to be stored in cache representing individual people
 */
interface CachePresence extends CacheItem {
  /**
   * json representing a person stored as string
   */
  presence?: string;
}

/**
 * Defines the expiration time
 */
const getPresenceInvalidationTime = (): number =>
  CacheService.config.presence.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether the groups store is enabled
 */
const presenceCacheEnabled = (): boolean => CacheService.config.presence.isEnabled && CacheService.config.isEnabled;

/**
 * async promise, allows developer to get my presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getMyPresence(graph: IGraph): Promise<Presence> {
  const betaGraph = BetaGraph.fromGraph(graph);
  const scopes = 'presence.read';
  let cache: CacheStore<CachePresence>;

  if (presenceCacheEnabled()) {
    cache = CacheService.getCache(cacheSchema, 'presence');
    const myPresence = await cache.getValue('me');
    if (myPresence && getPresenceInvalidationTime() > Date.now() - myPresence.timeCached) {
      return JSON.parse(myPresence.presence);
    }
  }

  const result = await betaGraph
    .api('/me/presence')
    .middlewareOptions(prepScopes(scopes))
    .get();

  if (presenceCacheEnabled()) {
    cache.putValue('me', { presence: result });
  }

  return result;
}

/**
 * async promise, allows developer to get user presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getUserPresence(graph: IGraph, userId: string): Promise<Presence> {
  const betaGraph = BetaGraph.fromGraph(graph);
  const scopes = ['presence.read', 'presence.read.all'];
  let cache: CacheStore<CachePresence>;

  if (presenceCacheEnabled()) {
    cache = CacheService.getCache(cacheSchema, 'presence');
    const myPresence = await cache.getValue(userId);
    if (myPresence && getPresenceInvalidationTime() > Date.now() - myPresence.timeCached) {
      return JSON.parse(myPresence.presence);
    }
  }

  const result = await betaGraph
    .api(`/users/${userId}/presence`)
    .middlewareOptions(prepScopes(...scopes))
    .get();
  if (presenceCacheEnabled()) {
    cache.putValue(userId, { presence: JSON.stringify(result) });
  }

  return result;
}

/**
 * async promise, allows developer to get person presense by providing array of IDynamicPerson
 *
 * @returns {}
 * @memberof BetaGraph
 */
export async function getUsersPresenceByPeople(graph: IGraph, people?: IDynamicPerson[]) {
  if (!people || people.length === 0) {
    return {};
  }

  const betaGraph = BetaGraph.fromGraph(graph);
  const batch = betaGraph.createBatch();
  const peoplePresence = {};
  let cache: CacheStore<CachePresence>;

  if (presenceCacheEnabled()) {
    cache = CacheService.getCache(cacheSchema, 'presence');
  }

  for (const person of people) {
    if (person !== '' && person.id) {
      const id = person.id;
      if (presenceCacheEnabled()) {
        const presence = await cache.getValue(id);
        if (presence && getPresenceInvalidationTime() > Date.now() - (await presence).timeCached) {
          peoplePresence[id] = JSON.parse(presence.presence);
        }
      } else {
        batch.get(id, `/users/${id}/presence`, ['presence.read', 'presence.read.all']);
      }
    }
  }

  try {
    const response = await batch.executeAll();

    for (const r of response.values()) {
      peoplePresence[r.id] = r.content;
      cache.putValue(r.id, { presence: JSON.stringify(r.content) });
    }

    return peoplePresence;
  } catch (_) {
    try {
      const response = await Promise.all(
        people.filter(person => person && person.id).map(person => getUserPresence(betaGraph, person.id))
      );

      for (const r of response) {
        peoplePresence[r.id] = r;
        cache.putValue(r.id, { presence: JSON.stringify(r) });
      }
      return peoplePresence;
    } catch (_) {
      return null;
    }
  }
}
