/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Message } from '@microsoft/microsoft-graph-types';

/**
 * Get common message with a user
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @param {string} emailAddress
 * @returns {Promise<IMessage[]>}
 */
export async function getMessagesWithUser(graph: IGraph, emailAddress: string): Promise<Message[]> {
  const response = await graph
    .api('/me/messages')
    .search(`"from:${emailAddress}"`)
    .get();
  return response.value;
}
