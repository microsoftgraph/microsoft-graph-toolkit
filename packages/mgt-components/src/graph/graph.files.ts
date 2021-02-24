/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Graph, IGraph, prepScopes } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { group } from 'console';

export async function getDriveItemByQuery(graph: IGraph, resource: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(resource)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}

  return response || null;
}

// GET /drives/{drive-id}/items/{item-id}
export async function getDriveItemById(graph: IGraph, driveId: string, itemId: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/drives/${driveId}/items/${itemId}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /drives/{drive-id}/root:/{item-path}
export async function getDriveItemByPath(graph: IGraph, driveId: string, itemPath: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/drives/${driveId}/root:/${itemPath}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /groups/{group-id}/drive/items/{item-id}
export async function getGroupDriveItemById(graph: IGraph, groupId: string, itemId: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/groups/${groupId}/drive/items/${itemId}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /groups/{group-id}/drive/root:/{item-path}
export async function getGroupDriveItemByPath(graph: IGraph, groupId: string, itemPath: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/groups/${groupId}/drive/root:/${itemPath}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /me/drive/items/{item-id}
export async function getMyDriveItemById(graph: IGraph, itemId: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/me/drive/items/${itemId}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /me/drive/root:/{item-path}
export async function getMyDriveItemByPath(graph: IGraph, itemPath: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/me/drive/root:/${itemPath}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /sites/{site-id}/drive/items/{item-id}
export async function getSiteDriveItemById(graph: IGraph, siteId: string, itemId: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/sites/${siteId}/drive/items/${itemId}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /sites/{site-id}/drive/root:/{item-path}
export async function getSiteDriveItemByPath(graph: IGraph, siteId: string, itemPath: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/sites/${siteId}/drive/root:/${itemPath}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /sites/{site-id}/lists/{list-id}/items/{item-id}/driveItem
export async function getListDriveItemById(
  graph: IGraph,
  siteId: string,
  listId: string,
  itemId: string
): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /users/{user-id}/drive/items/{item-id}
export async function getUserDriveItemById(graph: IGraph, userId: string, itemId: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/drive/items/${itemId}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /users/{user-id}/drive/root:/{item-path}
export async function getUserDriveItemByPath(graph: IGraph, userId: string, itemPath: string): Promise<DriveItem> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/drive/root:/${itemPath}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /me/insights/trending/{id}/resource
// GET /me/insights/used/{id}/resource
// GET /me/insights/shared/{id}/resource
export async function getMyInsightsDriveItemById(graph: IGraph, insightType: string, id: string): Promise<DriveItem> {
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph
      .api(`/me/insights/${insightType}/${id}/resource`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /users/{id or userPrincipalName}/insights/{trending or used or shared}/{id}/resource
export async function getUserInsightsDriveItemById(
  graph: IGraph,
  userId: string,
  insightType: string,
  id: string
): Promise<DriveItem> {
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/insights/${insightType}/${id}/resource`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}

// GET /me/drive/root/children
export async function getFiles(graph: IGraph): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api('/me/drive/root/children')
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /drives/{drive-id}/items/{item-id}/children
export async function getDriveFilesById(graph: IGraph, driveId: string, itemId: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/drives/${driveId}/items/${itemId}/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /drives/{drive-id}/root:/{path-relative-to-root}:/children
export async function getDriveFilesByPath(graph: IGraph, driveId: string, itemPath: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/drives/${driveId}/root:/${itemPath}:/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /groups/{group-id}/drive/items/{item-id}/children
export async function getGroupFilesById(graph: IGraph, groupId: string, itemId: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/groups/${groupId}/drive/items/${itemId}/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /me/drive/items/{item-id}/children
export async function getFilesById(graph: IGraph, itemId: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/me/drive/items/${itemId}/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /sites/{site-id}/drive/items/{item-id}/children
export async function getSiteFilesById(graph: IGraph, siteId: string, itemId: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/sites/${siteId}/drive/items/${itemId}/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /users/{user-id}/drive/items/{item-id}/children
export async function getUserFilesById(graph: IGraph, userId: string, itemId: string): Promise<DriveItem[]> {
  const scopes = 'files.read';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/drive/items/${itemId}/children`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /me/insights/{trending	| used | shared}
export async function getMyInsightsFiles(graph: IGraph, insightType: string): Promise<DriveItem[]> {
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph
      .api(`/me/insights/${insightType}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

// GET /users/{id | userPrincipalName}/insights/{trending	| used | shared}
export async function getUserInsightsFiles(graph: IGraph, userId: string, insightType: string): Promise<DriveItem[]> {
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/insights/${insightType}`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response.value || null;
}

export async function getFilesByListQuery(graph: IGraph, listQuery: string): Promise<DriveItem[]> {
  const scopes = ['files.read', 'sites.read.all'];
  let response;
  try {
    response = await graph
      .api(listQuery)
      .middlewareOptions(prepScopes(...scopes))
      .get();
  } catch {}
  return response.value || null;
}

export async function getFilesByQueries(graph: IGraph, fileQueries: string[]): Promise<DriveItem[]> {
  if (!fileQueries || fileQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch();
  const files = [];
  const scopes = ['files.read'];

  for (const fileQuery of fileQueries) {
    if (fileQuery !== '') {
      batch.get(fileQuery, fileQuery, scopes);
    }
  }

  try {
    const responses = await batch.executeAll();

    for (const fileQuery of fileQueries) {
      const response = responses.get(fileQuery);
      if (response && response.content && response.content.value && response.content.value.length > 0) {
        files.push(response.content.value[0]);
      }
    }

    return files;
  } catch (_) {
    try {
      return Promise.all(
        fileQueries
          .filter(fileQuery => fileQuery && fileQuery !== '')
          .map(async fileQuery => {
            const file = await getDriveItemByQuery(graph, fileQuery);
            if (file) {
              return file;
            }
          })
      );
    } catch (_) {
      return [];
    }
  }
}
