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
    const presence = await cache.getValue('me');
    if (presence && getPresenceInvalidationTime() > Date.now() - presence.timeCached) {
      return JSON.parse(presence.presence);
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
    const presence = await cache.getValue(userId);
    if (presence && getPresenceInvalidationTime() > Date.now() - presence.timeCached) {
      return JSON.parse(presence.presence);
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
  const peoplePresence = {};
  const peoplePresenceToQuery: string[] = [];
  const scopes = ['presence.read.all'];
  let cache: CacheStore<CachePresence>;

  if (presenceCacheEnabled()) {
    cache = CacheService.getCache(cacheSchema, 'presence');
  }

  for (const person of people) {
    if (person !== '' && person.id) {
      const id = person.id;
      peoplePresence[id] = null;
      let presence: CachePresence;
      if (presenceCacheEnabled()) {
        presence = await cache.getValue(id);
      }
      if (
        presenceCacheEnabled() &&
        presence &&
        getPresenceInvalidationTime() > Date.now() - (await presence).timeCached
      ) {
        peoplePresence[id] = JSON.parse(presence.presence);
      } else {
        peoplePresenceToQuery.push(id);
      }
    }
  }

  try {
    if (peoplePresenceToQuery.length > 0) {
      const presenceResult = await betaGraph
        .api(`/communications/getPresencesByUserId`)
        .middlewareOptions(prepScopes(...scopes))
        .post({
          ids: peoplePresenceToQuery
        });

      for (const r of presenceResult.value) {
        peoplePresence[r.id] = r;
        if (presenceCacheEnabled()) {
          cache.putValue(r.id, { presence: JSON.stringify(r) });
        }
      }
    }

    return peoplePresence;
  } catch (_) {
    try {
      /**
       * individual calls to getUserPresence as fallback
       * must filter out the contacts, which will either 404 or have PresenceUnknown response
       * caching will be handled by getUserPresence
       */
      const response = await Promise.all(
        people
          .filter(
            person =>
              person &&
              person.id &&
              !peoplePresence[person.id] &&
              'personType' in person &&
              (person as any).personType.subclass === 'OrganizationUser'
          )
          .map(person => getUserPresence(betaGraph, person.id))
      );

      for (const r of response) {
        peoplePresence[r.id] = r;
      }
      return peoplePresence;
    } catch (_) {
      return null;
    }
  }
}
