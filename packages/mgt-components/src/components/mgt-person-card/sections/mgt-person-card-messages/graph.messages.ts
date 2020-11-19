/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Message } from '@microsoft/microsoft-graph-types';

/**
 * Display metadata for a message item
 */
export interface IMessage extends Message {}

/**
 * Get messages for a user
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {IMessage}
 */
export async function getMessages(graph: IGraph): Promise<IMessage[]> {
  const response = await graph.api(`/me/messages`).get();
  return response.value;
}

/**
 * Get common message with a user
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @param {string} emailAddress
 * @returns {Promise<IMessage[]>}
 */
export async function getMessagesWithUser(graph: IGraph, emailAddress: string): Promise<IMessage[]> {
  const response = await graph.api(`/me/messages?$search="participants:${emailAddress}"`).get();
  return response.value;
}
