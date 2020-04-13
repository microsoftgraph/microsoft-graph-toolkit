/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IDynamicPerson } from '../components/mgt-person/mgt-person';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';
import { blobToBase64 } from '../utils/Utils';
import { findUserByEmail, getEmailFromGraphEntity } from './graph.people';

/**
 * retrieves a photo for the specified resource.
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
async function getPhotoForResource(graph: IGraph, resource: string, scopes: string[]): Promise<string> {
  try {
    const blob = await graph
      .api(`${resource}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .middlewareOptions(prepScopes(...scopes))
      .get();
    return await blobToBase64(blob);
  } catch (e) {
    return null;
  }
}

/**
 * async promise, returns Graph photos associated with contacts of the logged in user
 * @param contactId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function getContactPhoto(graph: IGraph, contactId: string): Promise<string> {
  return getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
}

/**
 * async promise, returns Graph photo associated with provided userId
 * @param userId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function getUserPhoto(graph: IGraph, userId: string): Promise<string> {
  return getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
}

/**
 * async promise, returns Graph photo associated with the logged in user
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function myPhoto(graph: IGraph): Promise<string> {
  return getPhotoForResource(graph, 'me', ['user.read']);
}

/**
 * async promise, loads image of user
 *
 * @export
 */
export async function getPersonImage(graph: IGraph, person: IDynamicPerson) {
  let image: string;

  if ((person as MicrosoftGraph.Person).userPrincipalName) {
    // try to find a user by userPrincipalName
    const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
    image = await getUserPhoto(graph, userPrincipalName);
  } else {
    // try to find a user by e-mail
    const email = getEmailFromGraphEntity(person);
    if (email) {
      const users = await findUserByEmail(graph, email);
      if (users && users.length) {
        // Check for an OrganizationUser
        const orgUser = users.find(p => {
          return (p as any).personType && (p as any).personType.subclass === 'OrganizationUser';
        });
        if (orgUser) {
          // Lookup by userId
          const userId = (users[0] as MicrosoftGraph.Person).scoredEmailAddresses[0].address;
          image = await getUserPhoto(graph, userId);
        } else {
          // Lookup by contactId
          const contactId = users[0].id;
          image = await getContactPhoto(graph, contactId);
        }
      }
    }
  }

  return image;
}
