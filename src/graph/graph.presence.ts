/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Presence } from '@microsoft/microsoft-graph-types-beta';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * async promise, allows developer to get my presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getMyPresence(graph: IGraph): Promise<Presence> {
  const scopes = 'presence.read';
  const result = await graph
    .api('/me/presence')
    .middlewareOptions(prepScopes(scopes))
    .get();

  return result;
}

/**
 * async promise, allows developer to get person presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getPersonPresence(graph: IGraph, userId: string): Promise<Presence> {
  const scopes = 'presence.read, presence.read.all';
  const result = await graph
    .api(`/users/${userId}/presence`)
    .middlewareOptions(prepScopes(scopes))
    .get();

  return result;
}

/**
 * async promise, allows developer to get person presense
 *
 * @returns {}
 * @memberof BetaGraph
 */
export async function getPeoplePresence(graph: IGraph, userIds: string[]) {
  if (!userIds || userIds.length === 0) {
    return [];
  }

  const batch = graph.createBatch();

  for (const id of userIds) {
    if (id !== '') {
      batch.get(id, `/users/${id}/presence`, ['presence.read', 'presence.read.all']);
    }
  }

  try {
    const response = await batch.execute();
    const peoplePresence = [];

    for (const id of userIds) {
      // tslint:disable
      console.log('RESPONSE: ' + JSON.stringify(response));
      const personPresence = response[id];
      if (personPresence && personPresence.id) {
        peoplePresence.push(personPresence);
        // { id: id, presence: personPresence }
      }
    }
    return peoplePresence;
  } catch (_) {
    try {
      return Promise.all(userIds.filter(id => id && id !== '').map(id => getPersonPresence(graph, id)));
    } catch (_) {
      return [];
    }
  }
}
