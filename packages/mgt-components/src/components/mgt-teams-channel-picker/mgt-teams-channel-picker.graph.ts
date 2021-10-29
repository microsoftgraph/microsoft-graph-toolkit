/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, BetaGraph, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { Team } from '@microsoft/microsoft-graph-types';
import { getPhotoForResource, CachePhoto } from '../../graph/graph.photos';
import { blobToBase64 } from '../../utils/Utils';
import { ResponseType } from '@microsoft/microsoft-graph-client';

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

export async function getTeamsPhotosforPhotoIds(graph: BetaGraph, teamIds: string[]): Promise<any> {
  let cache: CacheStore<CachePhoto>;
  let photoDetails: CachePhoto;

  // const batch = graph.createBatch();
  let scopes = ['team.readbasic.all'];
  let teamDict = {};

  for (const id of teamIds) {
    try {
      teamDict[id] = await getPhotoForResource(graph, `/teams/${id}`, scopes);
    } catch (_) {}
  }

  return teamDict;
}
