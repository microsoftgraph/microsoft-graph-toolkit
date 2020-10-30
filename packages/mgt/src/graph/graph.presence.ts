/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Presence } from '@microsoft/microsoft-graph-types-beta';
import { BetaGraph } from '../BetaGraph';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';
import { IDynamicPerson } from './types';

/**
 * async promise, allows developer to get user presense
 *
 * @returns {Promise<Presence>}
 * @param {IGraph} graph
 * @param {string} userId - id for the user or null for current signed in user
 * @memberof BetaGraph
 */
export async function getUserPresence(graph: IGraph, userId?: string): Promise<Presence> {
  const betaGraph = BetaGraph.fromGraph(graph);

  const scopes = userId ? ['presence.read.all'] : ['presence.read'];
  const resource = userId ? `/users/${userId}/presence` : '/me/presence';

  const result = await betaGraph
    .api(resource)
    .middlewareOptions(prepScopes(...scopes))
    .get();

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

  for (const person of people) {
    if (person !== '' && person.id) {
      const id = person.id;
      batch.get(id, `/users/${id}/presence`, ['presence.read', 'presence.read.all']);
    }
  }

  try {
    const peoplePresence = {};
    const response = await batch.executeAll();

    for (const r of response.values()) {
      peoplePresence[r.id] = r.content;
    }

    return peoplePresence;
  } catch (_) {
    try {
      const peoplePresence = {};
      const response = await Promise.all(
        people.filter(person => person && person.id).map(person => getUserPresence(betaGraph, person.id))
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
