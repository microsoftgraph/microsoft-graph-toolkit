/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '../../../../../IGraph';

/**
 * foo
 */
export interface IMessage {
  // tslint:disable-next-line: completed-docs
  receivedDateTime: Date;
  // tslint:disable-next-line: completed-docs
  subject: string;
  // tslint:disable-next-line: completed-docs
  from: { emailAddress: { address: string; name: string } };
  // tslint:disable-next-line: completed-docs
  bodyPreview: string;
}

/**
 * foo
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {IMessage}
 */
export async function getMessages(graph: IGraph, userId: string): Promise<IMessage[]> {
  const response = await graph.api(`/users/${userId}/messages`).get();
  return response.value;
}

/**
 * foo
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @param {string} emailAddress
 * @returns {Promise<IMessage[]>}
 */
export async function getMessagesWithUser(graph: IGraph, userId: string, emailAddress: string): Promise<IMessage[]> {
  const response = await graph
    .api(`/users/${userId}/messages?$filter=(from/emailAddress/address) eq '${emailAddress}'`)
    .get();
  return response.value;
}
