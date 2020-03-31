/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export async function getMe(graph: IGraph): Promise<User> {
  const batch = graph.createBatch();
  batch.get('user', 'me', ['user.read']);

  try {
    const response = await batch.execute();
    return response.user;
  } catch {
    return null;
  }
}

/**
 * async promise, returns all Graph users associated with the userPrincipleName provided
 *
 * @param {string} userPrincipleName
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export function getUser(graph: IGraph, userPrincipleName: string): Promise<User> {
  const scopes = 'user.readbasic.all';
  return graph
    .api(`/users/${userPrincipleName}`)
    .middlewareOptions(prepScopes(scopes))
    .get();
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
    const response = await batch.execute();
    const people = [];

    // iterate over userIds to ensure the order of ids
    for (const id of userIds) {
      const person = response[id];
      if (person && person.id) {
        people.push(person);
      }
    }

    return people;
  } catch {
    // fallback to making the request one by one
    try {
      return Promise.all(userIds.filter(id => id && id !== '').map(id => getUser(graph, id)));
    } catch {
      return [];
    }
  }
}
