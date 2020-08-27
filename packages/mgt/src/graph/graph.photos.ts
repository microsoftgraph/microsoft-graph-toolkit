/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { CacheItem, CacheSchema, CacheService, CacheStore } from '../utils/Cache';
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
export const getPhotoInvalidationTime = () =>
  CacheService.config.photos.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether photo store is enabled
 */
export const photosCacheEnabled = () => CacheService.config.photos.isEnabled && CacheService.config.isEnabled;

/**
 * Name of the users store name
 */
const userStore: string = 'users';

/**
 * Name of the contacts store name
 */
const contactStore: string = 'contacts';

/**
 * retrieves a photo for the specified resource.
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
export async function getPhotoForResource(graph: IGraph, resource: string, scopes: string[]): Promise<CachePhoto> {
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
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;
  if (photosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(cacheSchema, contactStore);
    photoDetails = await cache.getValue(contactId);
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    }
  }

  photoDetails = await getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
  if (photosCacheEnabled() && photoDetails) {
    cache.putValue(contactId, photoDetails);
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
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;
  if (photosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(cacheSchema, userStore);
    photoDetails = await cache.getValue(userId);
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    }
  }

  photoDetails = await getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
  if (photosCacheEnabled() && photoDetails) {
    cache.putValue(userId, photoDetails);
  }
  return photoDetails ? photoDetails.photo : null;
}

/**
 * async promise, returns Graph photo associated with the logged in user
 * @returns {Promise<string>}
 * @memberof Graph
 */
export async function myPhoto(graph: IGraph): Promise<string> {
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;
  if (photosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(cacheSchema, userStore);
    photoDetails = await cache.getValue('me');
    if (photoDetails && getPhotoInvalidationTime() > Date.now() - photoDetails.timeCached) {
      return photoDetails.photo;
    }
  }

  photoDetails = await getPhotoForResource(graph, 'me', ['user.read']);
  if (photosCacheEnabled()) {
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
  } else if ('personType' in person && (person as any).personType.subclass === 'PersonalContact') {
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

/**
 * checks if user has a photo in the cache
 * @param userId
 * @returns {CachePhoto}
 * @memberof Graph
 */
export async function getPhotoFromCache(userId: string, storeName: string): Promise<CachePhoto> {
  const cache = CacheService.getCache<CachePhoto>(cacheSchema, storeName);
  const item = await cache.getValue(userId);
  return item;
}

/**
 * checks if user has a photo in the cache
 * @param userId
 * @returns {void}
 * @memberof Graph
 */
export async function storePhotoInCache(userId: string, storeName: string, value: CachePhoto): Promise<void> {
  const cache = CacheService.getCache<CachePhoto>(cacheSchema, storeName);
  cache.putValue(userId, value);
}
