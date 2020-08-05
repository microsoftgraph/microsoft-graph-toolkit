/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { CacheItem, CacheSchema, CacheService } from '../utils/Cache';
import { prepScopes } from '../utils/GraphHelpers';
import { blobToBase64 } from '../utils/Utils';
import { findContactByEmail, findUserByEmail, getEmailFromGraphEntity } from './graph.people';
import { IDynamicPerson } from './types';

/**
 * defines the structure of the cache
 */
const cacheSchema: CacheSchema = {
  name: 'photos',
  stores: {
    contacts: {},
    users: {}
  },
  version: 1
};

/**
 * photo object stored in cache
 */
interface CachePhoto extends CacheItem {
  /**
   * user tag associated with photo
   */
  eTag?: string;
  /**
   * user/contact photo
   */
  photo?: string;
}

/**
 * Defines expiration time
 */
const photoInvalidationTime = CacheService.config.photos.invalidiationPeriod
  ? CacheService.config.photos.invalidiationPeriod
  : CacheService.config.defaultInvalidationPeriod;

/**
 * Whether photo store is enabled
 */
const photosCacheEnabled = CacheService.config.photos.isEnabled;

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
  const cache = CacheService.getCache<CachePhoto>(cacheSchema, 'contacts');

  // try to get a photo from cache
  let photoDetails: CachePhoto = photosCacheEnabled ? await cache.getValue(contactId) : null;

  // if photo is cached, use it
  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > photoInvalidationTime) {
      // check if new image is available and update for next time
      graph //
        .api(`me/contacts/${contactId}/photo`)
        .get()
        .then(
          async response => {
            // no eTag for contacts :(
            if (response) {
              photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
            }
            if (photosCacheEnabled) {
              cache.putValue(contactId, photoDetails);
            }
          },
          error => {
            // just add an empty photo to cache
            if (photosCacheEnabled) {
              cache.putValue(contactId, {});
            }
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
  if (photosCacheEnabled) {
    cache.putValue(contactId, photoDetails || {});
  }
  return photoDetails ? photoDetails.photo : null;
}

/**
 * async promise, returns Graph photo associated with provided userId
 * @param userId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export async function getUserPhoto(graph: IGraph, userId: string): Promise<string> {
  const cache = CacheService.getCache<CachePhoto>(cacheSchema, 'users');
  let photoDetails: CachePhoto = photosCacheEnabled ? await cache.getValue(userId) : null;

  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > photoInvalidationTime) {
      // check if new image is available and update for next time
      graph
        .api(`users/${userId}/photo`)
        .get()
        .then(
          async response => {
            if (response && response['@odata.mediaEtag'] !== photoDetails.eTag) {
              photoDetails = await getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
            }
            if (photosCacheEnabled) {
              cache.putValue(userId, photoDetails);
            }
          },
          error => {
            if (photosCacheEnabled) {
              cache.putValue(userId, {});
            }
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
  if (photosCacheEnabled && photoDetails) {
    cache.putValue(userId, photoDetails || {});
  }

  return photoDetails ? photoDetails.photo : null;
}

/**
 * async promise, returns Graph photo associated with the logged in user
 * @returns {Promise<string>}
 * @memberof Graph
 */
export async function myPhoto(graph: IGraph): Promise<string> {
  const cache = CacheService.getCache<CachePhoto>(cacheSchema, 'users');

  let photoDetails: CachePhoto = photosCacheEnabled ? await cache.getValue('me') : null;
  if (photoDetails) {
    if (Date.now() - photoDetails.timeCached > photoInvalidationTime) {
      // check if new image is available and update for next time
      graph
        .api('me/photo')
        .get()
        .then(
          async response => {
            if (response && response['@odata.mediaEtag'] !== photoDetails.eTag) {
              photoDetails = await getPhotoForResource(graph, 'me', ['user.read']);
            }
            if (photosCacheEnabled) {
              cache.putValue('me', photoDetails);
            }
          },
          error => {
            if (photosCacheEnabled) {
              cache.putValue('me', {});
            }
          }
        );
    }
    return photoDetails.photo;
  }

  photoDetails = await getPhotoForResource(graph, 'me', ['user.read']);
  if (photosCacheEnabled) {
    cache.putValue('me', photoDetails || {});
  }

  return photoDetails ? photoDetails.photo : null;
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
