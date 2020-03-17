/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Contact, Person, User } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * async promise, returns all Graph people who are most relevant contacts to the signed in user.
 *
 * @param {string} query
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function findPerson(graph: IGraph, query: string): Promise<Person[]> {
  const scopes = 'people.read';
  const result = await graph
    .api('/me/people')
    .search('"' + query + '"')
    .middlewareOptions(prepScopes(scopes))
    .get();
  return result ? result.value : null;
}

/**
 * async promise to the Graph for People, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeople(graph: IGraph): Promise<Person[]> {
  const scopes = 'people.read';

  const uri = '/me/people';
  const people = await graph
    .api(uri)
    .middlewareOptions(prepScopes(scopes))
    .filter("personType/class eq 'Person'")
    .get();
  return people ? people.value : null;
}

/**
 * async promise to the Graph for People, defined by a group id
 *
 * @param {string} groupId
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getPeopleFromGroup(graph: IGraph, groupId: string): Promise<Person[]> {
  const scopes = 'people.read';

  const uri = `/groups/${groupId}/members`;
  const people = await graph
    .api(uri)
    .middlewareOptions(prepScopes(scopes))
    .get();
  return people ? people.value : null;
}

/**
 * returns a promise that resolves after specified time
 * @param time in milliseconds
 */
export function getEmailFromGraphEntity(entity: User | Person | Contact): string {
  const person = entity as Person;
  const user = entity as User;
  const contact = entity as Contact;

  if (user.mail) {
    return user.mail;
  } else if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
    return person.scoredEmailAddresses[0].address;
  } else if (contact.emailAddresses && contact.emailAddresses.length) {
    return contact.emailAddresses[0].address;
  }

  return null;
}
