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
import { Photo } from '@microsoft/microsoft-graph-types';
import { isGraphError } from './isGraphError';

/**
 * async promise, returns IDynamicPerson
 *
 * @param {string} userId
 * @returns {(Promise<IDynamicPerson>)}
 * @memberof Graph
 */
export const getUserWithPhoto = async (
  graph: IGraph,
  userId?: string,
  requestedProps?: string[]
): Promise<IDynamicPerson> => {
  let photo: string;
  let user: IDynamicPerson = null;

  let cachedPhoto: CachePhoto;
  let cachedUser: CacheUser;

  const resource = userId ? `users/${userId}` : 'me';
  const fullResource = resource + (requestedProps ? `?$select=${requestedProps.toString()}` : '');

  const scopes = userId ? ['user.readbasic.all'] : ['user.read'];

  // attempt to get user and photo from cache if enabled
  if (getIsUsersCacheEnabled()) {
    const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
    cachedUser = await cache.getValue(userId || 'me');
    if (cachedUser && getUserInvalidationTime() > Date.now() - cachedUser.timeCached) {
      user = cachedUser.user ? (JSON.parse(cachedUser.user) as IDynamicPerson) : null;
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
        const response: Photo = (await graph.api(`${resource}/photo`).get()) as Photo;
        if (response?.['@odata.mediaEtag'] && response['@odata.mediaEtag'] === cachedPhoto.eTag) {
          // put current image into the cache to update the timestamp since etag is the same
          await storePhotoInCache(userId || 'me', schemas.photos.stores.users, cachedPhoto);
          photo = cachedPhoto.photo;
        } else {
          cachedPhoto = null;
        }
      } catch (e: unknown) {
        if (isGraphError(e)) {
          // if 404 received (photo not found) but user already in cache, update timeCache value to prevent repeated 404 error / graph calls on each page refresh
          // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
          if (e.code === 'ErrorItemNotFound' || e.code === 'ImageNotFound') {
            await storePhotoInCache(userId || 'me', schemas.photos.stores.users, { eTag: null, photo: null });
          }
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
      // eslint-disable-next-line @typescript-eslint/dot-notation
      eTag = photoResponse.headers['ETag'];
      photo = photoResponse.content as string;
    }

    const userResponse = response.get('user');
    if (userResponse) {
      user = userResponse.content as IDynamicPerson;
    }

    // store user & photo in their respective cache
    if (getIsUsersCacheEnabled()) {
      const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
      await cache.putValue(userId || 'me', { user: JSON.stringify(user) });
    }
    if (getIsPhotosCacheEnabled()) {
      await storePhotoInCache(userId || 'me', schemas.photos.stores.users, { eTag, photo });
    }
  } else if (!cachedPhoto) {
    try {
      // if only photo or user is not cached, get it individually
      const response = await getPhotoForResource(graph, resource, scopes);
      if (response) {
        if (getIsPhotosCacheEnabled()) {
          await storePhotoInCache(userId || 'me', schemas.photos.stores.users, {
            eTag: response.eTag,
            photo: response.photo
          });
        }
        photo = response.photo;
      }
    } catch (_) {
      // intentionally left empty...
    }
  } else if (!cachedUser) {
    // get user from graph
    try {
      const response: IDynamicPerson = (await graph
        .api(fullResource)
        .middlewareOptions(prepScopes(...scopes))
        .get()) as IDynamicPerson;

      if (response) {
        if (getIsUsersCacheEnabled()) {
          const cache = CacheService.getCache<CacheUser>(schemas.users, schemas.users.stores.users);
          await cache.putValue(userId || 'me', { user: JSON.stringify(response) });
        }
        user = response;
      }
    } catch (_) {
      // intentionally left empty...
    }
  }

  if (user) {
    user.personImage = photo;
  }
  return user;
};
