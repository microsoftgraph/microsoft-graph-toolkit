/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { DriveItem, SharedInsight, ThumbnailSet } from '@microsoft/microsoft-graph-types';

/**
 * Potential file icon types
 */
export enum IconType {
  Word,
  PowerPoint,
  Other
}

/**
 * Get files shared between me and another user.
 * TODO: Figure out the correct graph call.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {Promise<DriveItem[]>}
 */
export async function getFilesSharedByUser(graph: IGraph, emailAddress: string): Promise<SharedInsight[]> {
  // https://graph.microsoft.com/v1.0/me/insights/shared?$filter=lastshared/sharedby/address eq 'kellygraham@contoso.com'
  const response = await graph
    .api('me/insights/shared')
    .filter(`lastshared/sharedby/address eq '${emailAddress}'`)
    .get();
  return response.value || null;
}

export async function getMostRecentFiles(graph: IGraph): Promise<SharedInsight[]> {
  // https://graph.microsoft.com/v1.0/me/insights/shared?$filter=lastshared/sharedby/address eq 'kellygraham@contoso.com'
  const response = await graph.api('me/insights/used').get();
  return response.value || null;
}

export async function getThumbnails(graph: IGraph, item: DriveItem): Promise<ThumbnailSet[]> {
  const driveId = item.remoteItem.parentReference.driveId;
  const itemId = item.remoteItem.id;

  const response = await graph.api(`/drives/${driveId}/items/${itemId}/thumbnails`).get();
  return response.value || null;
}
