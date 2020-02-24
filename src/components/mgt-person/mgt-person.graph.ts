import { Contact } from '@microsoft/microsoft-graph-types';
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
