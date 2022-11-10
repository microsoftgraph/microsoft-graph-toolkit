/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheService, IGraph, prepScopes } from '@microsoft/mgt-element';

import {
  CachePhoto,
  getIsPhotosCacheEnabled,
  getPhotoForResource,
  getPhotoFromCache,
  getPhotoInvalidationTime,
  storePhotoInCache
} from './graph.photos';
import { CacheUser, getIsUsersCacheEnabled, getUserInvalidationTime } from './graph.user';
import { IDynamicPerson } from './types';
import { schemas } from './cacheStores';

/**
 * async promise, returns IDynamicPerson
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */
export async function getUserWithPhoto(
  graph: IGraph,
  userId?: string,
  requestedProps?: string[]
): Promise<IDynamicPerson> {
  let photo = null;
  let user: IDynamicPerson = null;

  let cachedPhoto: CachePhoto;
  let cachedUser: CacheUser;

  const resource = userId ? `users/${userId}` : 'me';
  let fullResource = resource + (requestedProps ? `?$select=${requestedProps.toString()}` : '');

  const scopes = userId ? ['user.readbasic.all'] : ['user.read'];

  // attempt to get user and photo from cache if enabled
  if (getIsUsersCacheEnabled()) {
    const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
    cachedUser = await cache.getValue(userId || 'me');
    if (cachedUser !== undefined && getUserInvalidationTime() > Date.now() - cachedUser.timeCached) {
      user = cachedUser.user ? JSON.parse(cachedUser.user) : null;
      if (user !== null && requestedProps) {
        const uniqueProps = requestedProps.filter(prop => !Object.keys(user).includes(prop));
        if (uniqueProps.length >= 1) {
          user = null;
          cachedUser = null;
        }
      }
    } else {
      cachedUser = null;
    }
  }
  if (getIsPhotosCacheEnabled()) {
    cachedPhoto = await getPhotoFromCache(userId || 'me', schemas.photos.stores.users);
    if (cachedPhoto && getPhotoInvalidationTime() > Date.now() - cachedPhoto.timeCached) {
      photo = cachedPhoto.photo;
    } else if (cachedPhoto) {
      try {
        const response = await graph.api(`${resource}/photo`).get();
        if (response && response['@odata.mediaEtag'] && response['@odata.mediaEtag'] === cachedPhoto.eTag) {
          // put current image into the cache to update the timestamp since etag is the same
          storePhotoInCache(userId || 'me', schemas.photos.stores.users, cachedPhoto);
          photo = cachedPhoto.photo;
        } else {
          cachedPhoto = null;
        }
      } catch (e) {
        //if 404 received (photo not found) but user already in cache, update timeCache value to prevent repeated 404 error / graph calls on each page refresh
        if (e.code === 'ErrorItemNotFound' || e.code === 'ImageNotFound') {
          storePhotoInCache(userId || 'me', schemas.photos.stores.users, { eTag: null, photo: null });
        }
      }
    }
  }

  // if both are not in the cache, batch get them
  if (!cachedPhoto && !cachedUser) {
    let eTag: string;

    // batch calls
    const batch = graph.createBatch();
    if (userId) {
      batch.get('user', `/users/${userId}${requestedProps ? '?$select=' + requestedProps.toString() : ''}`, [
        'user.readbasic.all'
      ]);
      batch.get('photo', `users/${userId}/photo/$value`, ['user.readbasic.all']);
    } else {
      batch.get('user', 'me', ['user.read']);
      batch.get('photo', 'me/photo/$value', ['user.read']);
    }
    const response = await batch.executeAll();

    const photoResponse = response.get('photo');
    if (photoResponse) {
      eTag = photoResponse.headers['ETag'];
      photo = photoResponse.content;
    }

    const userResponse = response.get('user');
    if (userResponse) {
      user = userResponse.content;
    }

    // store user & photo in their respective cache
    if (getIsUsersCacheEnabled()) {
      const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
      cache.putValue(userId || 'me', { user: JSON.stringify(user) });
    }
    if (getIsPhotosCacheEnabled()) {
      storePhotoInCache(userId || 'me', schemas.photos.stores.users, { eTag, photo: photo });
    }
  } else if (!cachedPhoto) {
    try {
      // if only photo or user is not cached, get it individually
      const response = await getPhotoForResource(graph, resource, scopes);
      if (response) {
        if (getIsPhotosCacheEnabled()) {
          storePhotoInCache(userId || 'me', schemas.photos.stores.users, {
            eTag: response.eTag,
            photo: response.photo
          });
        }
        photo = response.photo;
      }
    } catch (_) {}
  } else if (!cachedUser) {
    // get user from graph
    try {
      const response = await graph
        .api(fullResource)
        .middlewareOptions(prepScopes(...scopes))
        .get();

      if (response) {
        if (getIsUsersCacheEnabled()) {
          const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
          cache.putValue(userId || 'me', { user: JSON.stringify(response) });
        }
        user = response;
      }
    } catch (_) {}
  }

  if (user) {
    user.personImage = photo;
  }
  return user;
}
