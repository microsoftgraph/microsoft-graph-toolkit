/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';
import { findPerson } from './graph.people';
import { getPersonImage } from './graph.photos';
import { IDynamicPerson } from './types';

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export function getMe(graph: IGraph): Promise<User> {
  return graph
    .api('me')
    .middlewareOptions(prepScopes('user.read'))
    .get();
}

/**
 * asnyc promise, returns IDynamicPerson
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */

export async function getUserWithPhoto(graph: IGraph, userId?: string): Promise<IDynamicPerson> {
  const batch = graph.createBatch();
  let person = null as IDynamicPerson;

  if (userId) {
    batch.get('user', `/users/${userId}`, ['user.readbasic.all']);
    batch.get('photo', `users/${userId}/photo/$value`, ['user.readbasic.all']);
  } else {
    batch.get('user', 'me', ['user.read']);
    batch.get('photo', 'me/photo/$value', ['user.read']);
  }
  const response = await batch.execute();
  person = response.user;
  person.personImage = response.photo;

  return person;
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
  } catch (_) {
    // fallback to making the request one by one
    try {
      return Promise.all(userIds.filter(id => id && id !== '').map(id => getUser(graph, id)));
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

  for (const personQuery of peopleQueries) {
    if (personQuery !== '') {
      batch.get(personQuery, `/me/people?$search="${personQuery}"`, ['people.read']);
    }
  }

  try {
    const response = await batch.execute();

    for (const personQuery of peopleQueries) {
      const person = response[personQuery];
      if (person) {
        people.push(person.value[0]);
      }
    }

    return people;
  } catch (_) {
    try {
      for (const personQuery of peopleQueries) {
        if (personQuery !== '') {
          const person = (await findPerson(graph, personQuery)) as IDynamicPerson[];
          if (person && person.length) {
            people.push(person[0]);

            const image = await getPersonImage(graph, person[0]);
            if (image) {
              person[0].personImage = image;
            }
          }
        }
      }
      return people;
    } catch (_) {
      return [];
    }
  }
}
