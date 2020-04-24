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
 * Person Type enum
 *
 * @export
 * @enum {number}
 */
export enum PersonType {
  /**
   * Any type
   */
  Any = 0,

  /**
   * A Person such as User or Contact
   */
  // tslint:disable-next-line:no-bitwise
  Person = 1 << 0,

  /**
   * A group
   */
  // tslint:disable-next-line:no-bitwise
  Group = 1 << 1
}

/**
 * async promise, returns all Graph people who are most relevant contacts to the signed in user.
 *
 * @param {string} query
 * @param {number} [top=10] - number of people to return
 * @param {PersonType} [personType=PersonType.Person] - the type of person to search for
 * @returns {(Promise<Person[]>)}
 */
export async function findPeople(
  graph: IGraph,
  query: string,
  top: number = 10,
  personType: PersonType = PersonType.Person
): Promise<Person[]> {
  const scopes = 'people.read';

  let filterQuery = '';

  if (personType !== PersonType.Any) {
    filterQuery = `personType/class eq '${PersonType[personType]}'`;
  }

  const result = await graph
    .api('/me/people')
    .search('"' + query + '"')
    .top(top)
    .filter(filterQuery)
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

/**
 * async promise, returns a Graph contact associated with the email provided
 *
 * @param {string} email
 * @returns {(Promise<Contact[]>)}
 * @memberof Graph
 */
export async function findContactByEmail(graph: IGraph, email: string): Promise<Contact[]> {
  const scopes = 'contacts.read';
  const result = await graph
    .api('/me/contacts')
    .filter(`emailAddresses/any(a:a/address eq '${email}')`)
    .middlewareOptions(prepScopes(scopes))
    .get();
  return result ? result.value : null;
}

/**
 * async promise, returns Graph contact and/or Person associated with the email provided
 * Uses: Graph.findPerson(email) and Graph.findContactByEmail(email)
 *
 * @param {string} email
 * @returns {(Promise<Array<Person | Contact>>)}
 * @memberof Graph
 */
export function findUserByEmail(graph: IGraph, email: string): Promise<Array<Person | Contact>> {
  return Promise.all([findPeople(graph, email), findContactByEmail(graph, email)]).then(([people, contacts]) => {
    return ((people as any[]) || []).concat(contacts || []);
  });
}
