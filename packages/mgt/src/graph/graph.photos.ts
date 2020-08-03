/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';
import { blobToBase64 } from '../utils/Utils';
import { findContactByEmail, findUserByEmail, getEmailFromGraphEntity } from './graph.people';
import { IDynamicPerson } from './types';

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
  let email: string;

  if ((person as MicrosoftGraph.Person).userPrincipalName) {
    // try to find a user by userPrincipalName
    const userPrincipalName = (person as MicrosoftGraph.Person).userPrincipalName;
    image = await getUserPhoto(graph, userPrincipalName);
  } else if ((person as any).personType.subclass === 'PersonalContact') {
    // if person is a contact, look for them and their photo in contact api
    email = getEmailFromGraphEntity(person);
    const contact = await findContactByEmail(graph, email);
    if (contact && contact.length && contact[0].id) {
      image = await getContactPhoto(graph, contact[0].id);
    }
  } else if (person.id) {
    image = await getUserPhoto(graph, person.id);
  }
  if (image) {
    return image;
  }

  // try to find a user by e-mail
  email = getEmailFromGraphEntity(person);

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
        for (const user of users) {
          const contactId = user.id;
          image = await getContactPhoto(graph, contactId);
          if (image) {
            break;
          }
        }
      }
    }
  }

  return image;
}
