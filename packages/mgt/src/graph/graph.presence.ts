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
 * async promise, allows developer to get my presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getMyPresence(graph: IGraph): Promise<Presence> {
  const betaGraph = BetaGraph.fromGraph(graph);
  const scopes = 'presence.read';
  const result = await betaGraph
    .api('/me/presence')
    .middlewareOptions(prepScopes(scopes))
    .get();

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
  const result = await betaGraph
    .api(`/users/${userId}/presence`)
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
