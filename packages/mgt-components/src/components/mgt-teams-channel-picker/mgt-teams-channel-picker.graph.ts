/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, BetaGraph, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { Team } from '@microsoft/microsoft-graph-types';
import {
  getPhotoForResource,
  CachePhoto,
  getPhotoInvalidationTime,
  getIsPhotosCacheEnabled
} from '../../graph/graph.photos';
import { schemas } from '../../graph/cacheStores';

/**
 * async promise, returns all Teams associated with the user logged in
 *
 * @returns {Promise<Team[]>}
 * @memberof Graph
 */
export async function getAllMyTeams(graph: IGraph): Promise<Team[]> {
  const teams = await graph.api('/me/joinedTeams').select(['displayName', 'id', 'isArchived']).get();
  return teams ? teams.value : null;
}

/** An object collection of cached photos. */
type CachePhotos = {
  [key: string]: CachePhoto;
};

export async function getTeamsPhotosforPhotoIds(graph: BetaGraph, teamIds: string[]): Promise<CachePhotos> {
  let cache: CacheStore<CachePhoto>;
  let photos: CachePhotos = {};

  if (getIsPhotosCacheEnabled()) {
    cache = CacheService.getCache<CachePhoto>(schemas.photos, schemas.photos.stores.teams);
    for (const id of teamIds) {
      try {
        const photoDetail = await cache.getValue(id);
        if (photoDetail && getPhotoInvalidationTime() > Date.now() - photoDetail.timeCached) {
          photos[id] = photoDetail;
        }
      } catch (_) {}
    }
    if (Object.keys(photos).length) {
      return photos;
    }
  }

  let scopes = ['team.readbasic.all'];
  photos = {};

  for (const id of teamIds) {
    try {
      const photoDetail = await getPhotoForResource(graph, `/teams/${id}`, scopes);
      if (getIsPhotosCacheEnabled() && photoDetail) {
        cache.putValue(id, photoDetail);
      }
      photos[id] = photoDetail;
    } catch (_) {}
  }

  return photos;
}
