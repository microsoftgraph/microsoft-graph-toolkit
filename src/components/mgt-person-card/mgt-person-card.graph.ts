/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Presence } from '@microsoft/microsoft-graph-types-beta';
import { IGraph } from '../../IGraph';
import { prepScopes } from '../../utils/GraphHelpers';

/**
 * async promise, allows developer to get person presense
 *
 * @returns {Promise<Presence>}
 * @memberof BetaGraph
 */
export async function getMyPresence(graph: IGraph): Promise<Presence> {
  const scopes = 'presence.read';
  const result = await graph
    .api('/me/presence')
    .middlewareOptions(prepScopes(scopes))
    .get();

  return result;
}
