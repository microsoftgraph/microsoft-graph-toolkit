/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { Team } from '@microsoft/microsoft-graph-types';

/**
 * async promise, returns all Teams associated with the user logged in
 *
 * @returns {Promise<Team[]>}
 * @memberof Graph
 */
export async function getAllMyTeams(graph: IGraph): Promise<Team[]> {
  const scopes = 'team.readbasic.all';
  const teams = await graph
    .api('/me/joinedTeams')
    .select(['displayName', 'id', 'isArchived'])
    .middlewareOptions(prepScopes(scopes))
    .get();
  return teams ? teams.value : null;
}
