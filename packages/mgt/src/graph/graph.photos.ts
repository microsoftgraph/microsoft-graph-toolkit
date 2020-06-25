/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { Cache, CacheItem, CacheSchema } from '../utils/Cache';
import { prepScopes } from '../utils/GraphHelpers';
import { blobToBase64 } from '../utils/Utils';
import { findUserByEmail, getEmailFromGraphEntity } from './graph.people';
import { IDynamicPerson } from './types';

const cacheSchema: CacheSchema = {
  name: 'photos',
  stores: {
    contacts: {},
    users: {}
  },
  version: 1
};

interface CachePhoto extends CacheItem {
  eTag?: string;
  photo?: string;
}

// Time to invalidate cache in ms
// 3600000ms === 1hr
const cacheInvalidationTime = 3600000;

/**
 * retrieves a photo for the specified resource.
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
async function getPhotoForResource(graph: IGraph, resource: string, scopes: string[]): Promise<CachePhoto> {
  try {
    const response = (await graph
      .api(`${resource}/photo/$value`)
      .responseType(ResponseType.RAW)
      .middlewareOptions(prepScopes(...scopes))
      .get()) as Response;

    if (!response.ok) {
      return null;
    }

    const eTag = response.headers.get('ETag');
    const blob = await blobToBase64(await response.blob());
    return { eTag, photo: blob };
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
export async function getContactPhoto(graph: IGraph, contactId: string): Promise<string> {
  const cache = new Cache<CachePhoto>(cacheSchema, 'contacts');

  // try to get a photo from cache
  let photoDetails: CachePhoto = await cache.getValue(contactId);

  // if photo is cached, use it
  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > cacheInvalidationTime) {
      // check if new image is available and update for next time
      graph
        .api(`me/contacts/${contactId}/photo`)
        .get()
        .then(
          async response => {
            // no eTag for contacts :(
            if (response) {
              photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
            }
            cache.putValue(contactId, photoDetails);
          },
          error => {
            // just add an empty photo to cache
            cache.putValue(contactId, {});
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
  cache.putValue(contactId, photoDetails || {});
  return photoDetails ? photoDetails.photo : null;
}

/**
 * async promise, returns Graph photo associated with provided userId
 * @param userId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export async function getUserPhoto(graph: IGraph, userId: string): Promise<string> {
  const cache = new Cache<CachePhoto>(cacheSchema, 'users');

  let photoDetails: CachePhoto = await cache.getValue(userId);
  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > cacheInvalidationTime) {
      // check if new image is available and update for next time
      graph
        .api(`users/${userId}/photo`)
        .get()
        .then(
          async response => {
            if (response && response['@odata.mediaEtag'] !== photoDetails.eTag) {
              photoDetails = await getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
            }
            cache.putValue(userId, photoDetails);
          },
          error => {
            cache.putValue(userId, {});
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
  cache.putValue(userId, photoDetails || {});

  return photoDetails ? photoDetails.photo : null;
}

/**
 * async promise, returns Graph photo associated with the logged in user
 * @returns {Promise<string>}
 * @memberof Graph
 */
export async function myPhoto(graph: IGraph): Promise<string> {
  const cache = new Cache<CachePhoto>(cacheSchema, 'users');

  let photoDetails: CachePhoto = await cache.getValue('me');
  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > cacheInvalidationTime) {
      // check if new image is available and update for next time
      graph
        .api('me/photo')
        .get()
        .then(
          async response => {
            if (response && response['@odata.mediaEtag'] !== photoDetails.eTag) {
              photoDetails = await getPhotoForResource(graph, 'me', ['user.read']);
            }
            cache.putValue('me', photoDetails);
          },
          error => {
            cache.putValue('me', {});
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, 'me', ['user.read']);
  cache.putValue('me', photoDetails || {});

  return photoDetails ? photoDetails.photo : null;
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
    if (person.id) {
      image = await getUserPhoto(graph, person.id);
      if (image) {
        return image;
      }
    }

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
