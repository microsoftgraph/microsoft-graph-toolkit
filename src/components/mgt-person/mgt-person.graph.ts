import { Contact, Person } from '@microsoft/microsoft-graph-types';
import { findPerson } from '../../graph/graph.people';
import { IGraph } from '../../IGraph';
import { prepScopes } from '../../utils/GraphHelpers';

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
  return Promise.all([findPerson(graph, email), findContactByEmail(graph, email)]).then(([people, contacts]) => {
    return ((people as any[]) || []).concat(contacts || []);
  });
}
