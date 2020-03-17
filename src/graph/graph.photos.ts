/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ResponseType } from '@microsoft/microsoft-graph-client';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';
import { blobToBase64 } from '../utils/Utils';

/**
 * retrieves a photo for the specified resource.
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
async function getPhotoForResource(graph: IGraph, resource: string, scopes: string[]): Promise<string> {
  try {
    const blob = await graph
      .api(`${resource}/photo/$value`)
      .responseType(ResponseType.BLOB)
      .middlewareOptions(prepScopes(...scopes))
      .get();
    return await blobToBase64(blob);
  } catch (e) {
    return null;
  }
}

/**
 * async promise, returns Graph photos associated with contacts of the logged in user
 * @param contactId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function getContactPhoto(graph: IGraph, contactId: string): Promise<string> {
  return getPhotoForResource(graph, `me/contacts/${contactId}`, ['contacts.read']);
}

/**
 * async promise, returns Graph photo associated with provided userId
 * @param userId
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function getUserPhoto(graph: IGraph, userId: string): Promise<string> {
  return getPhotoForResource(graph, `users/${userId}`, ['user.readbasic.all']);
}

/**
 * async promise, returns Graph photo associated with the logged in user
 * @returns {Promise<string>}
 * @memberof Graph
 */
export function myPhoto(graph: IGraph): Promise<string> {
  return getPhotoForResource(graph, 'me', ['user.read']);
}
