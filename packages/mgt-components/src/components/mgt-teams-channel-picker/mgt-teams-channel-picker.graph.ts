/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, BetaGraph, CacheService, CacheStore, prepScopes } from '@microsoft/mgt-element';
import { Team } from '@microsoft/microsoft-graph-types';
import {
  getPhotoForResource,
  CachePhoto,
  getPhotoInvalidationTime,
  getIsPhotosCacheEnabled
} from '../../graph/graph.photos';
import { schemas } from '../../graph/cacheStores';
import { CollectionResponse } from '@microsoft/mgt-element';

/**
 * async promise, returns all Teams associated with the user logged in
 *
 * @returns {Promise<Team[]>}
 * @memberof Graph
 */
export const getAllMyTeams = async (graph: IGraph, scopes: string[]): Promise<Team[]> => {
  const teams = (await graph
    .api('/me/joinedTeams')
    .select(['displayName', 'id', 'isArchived'])
    .middlewareOptions(prepScopes(...scopes))
    .get()) as CollectionResponse<Team>;

  return teams?.value;
};

/** An object collection of cached photos. */
type CachePhotos = Record<string, CachePhoto>;

/**
 * Load the photos for a give set of teamIds
 *
 * @param graph {BetaGraph}
 * @param teamIds {string[]}
 * @returns {Promise<CachePhotos>}
 */
export const getTeamsPhotosforPhotoIds = async (graph: BetaGraph, teamIds: string[]): Promise<CachePhotos> => {
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
      } catch (_) {
        // no-op
      }
    }
    if (Object.keys(photos).length) {
      return photos;
    }
  }

  const scopes = ['team.readbasic.all'];
  photos = {};

  for (const id of teamIds) {
    try {
      const photoDetail = await getPhotoForResource(graph, `/teams/${id}`, scopes);
      if (getIsPhotosCacheEnabled() && photoDetail) {
        await cache.putValue(id, photoDetail);
      }
      photos[id] = photoDetail;
    } catch (_) {
      // no-op
    }
  }

  return photos;
};
