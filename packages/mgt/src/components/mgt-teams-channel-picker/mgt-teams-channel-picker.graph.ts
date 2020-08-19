/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Team } from '@microsoft/microsoft-graph-types';
import { prepScopes } from '../../utils/GraphHelpers';

/**
 * async promise, returns all Teams associated with the user logged in
 *
 * @returns {Promise<Team[]>}
 * @memberof Graph
 */
export async function getAllMyTeams(graph: IGraph): Promise<Team[]> {
  const teams = await graph
    .api('/me/joinedTeams')
    .middlewareOptions(prepScopes('User.Read.All'))
    .select(['displayName', 'id', 'isArchived'])
    .get();
  return teams ? teams.value : null;
}
