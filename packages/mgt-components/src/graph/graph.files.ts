/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphPageIterator, IGraph, prepScopes } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';

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
export async function getFilesIterator(graph: IGraph, top?: number): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;

  try {
    request = await graph.api('/me/drive/root/children').middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }

    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /drives/{drive-id}/items/{item-id}/children
export async function getDriveFilesByIdIterator(
  graph: IGraph,
  driveId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;

  try {
    request = graph.api(`/drives/${driveId}/items/${itemId}/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /drives/{drive-id}/root:/{item-path}:/children
export async function getDriveFilesByPathIterator(
  graph: IGraph,
  driveId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;

  try {
    request = graph.api(`/drives/${driveId}/root:/${itemPath}:/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /groups/{group-id}/drive/items/{item-id}/children
export async function getGroupFilesByIdIterator(
  graph: IGraph,
  groupId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;

  try {
    request = graph.api(`/groups/${groupId}/drive/items/${itemId}/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /groups/{group-id}/drive/root:/{item-path}:/children
export async function getGroupFilesByPathIterator(
  graph: IGraph,
  groupId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;

  try {
    request = graph.api(`/groups/${groupId}/drive/root:/${itemPath}:/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /me/drive/items/{item-id}/children
export async function getFilesByIdIterator(
  graph: IGraph,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/me/drive/items/${itemId}/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /me/drive/root:/{item-path}:/children
export async function getFilesByPathIterator(
  graph: IGraph,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/me/drive/root:/${itemPath}:/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /sites/{site-id}/drive/items/{item-id}/children
export async function getSiteFilesByIdIterator(
  graph: IGraph,
  siteId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/sites/${siteId}/drive/items/${itemId}/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /sites/{site-id}/drive/root:/{item-path}:/children
export async function getSiteFilesByPathIterator(
  graph: IGraph,
  siteId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/sites/${siteId}/drive/root:/${itemPath}:/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /users/{user-id}/drive/items/{item-id}/children
export async function getUserFilesByIdIterator(
  graph: IGraph,
  userId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/users/${userId}/drive/items/${itemId}/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /users/{user-id}/drive/root:/{item-path}:/children
export async function getUserFilesByPathIterator(
  graph: IGraph,
  userId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = 'files.read';
  let request;
  let filesPageIterator;
  try {
    request = graph.api(`/users/${userId}/drive/root:/${itemPath}:/children`).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

export async function getFilesByListQueryIterator(
  graph: IGraph,
  listQuery: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const scopes = ['files.read', 'sites.read.all'];
  let request;
  let filesPageIterator;

  try {
    request = await graph.api(listQuery).middlewareOptions(prepScopes(...scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIterator(graph, request);
  } catch {}
  return filesPageIterator || null;
}

// GET /me/insights/{trending	| used | shared}
export async function getMyInsightsFiles(graph: IGraph, insightType: string): Promise<DriveItem[]> {
  const scopes = ['sites.read.all'];
  let insightResponse;
  try {
    insightResponse = await graph
      .api(`/me/insights/${insightType}`)
      .filter(`resourceReference/type eq 'microsoft.graph.driveItem'`)
      .middlewareOptions(prepScopes(...scopes))
      .get();
  } catch {}

  const result = getDriveItemsByInsights(graph, insightResponse, scopes);
  return result || null;
}

// GET /users/{id | userPrincipalName}/insights/{trending	| used | shared}
export async function getUserInsightsFiles(graph: IGraph, userId: string, insightType: string): Promise<DriveItem[]> {
  const scopes = ['sites.read.all'];
  let insightResponse;
  let endpoint;
  let filter;

  if (insightType === 'shared') {
    endpoint = `/me/insights/shared`;
    filter = `((lastshared/sharedby/id eq '${userId}') and (resourceReference/type eq 'microsoft.graph.driveItem'))`;
  } else {
    endpoint = `/users/${userId}/insights/${insightType}`;
    filter = `resourceReference/type eq 'microsoft.graph.driveItem'`;
  }

  try {
    insightResponse = await graph
      .api(endpoint)
      .filter(filter)
      .middlewareOptions(prepScopes(...scopes))
      .get();
  } catch {}

  const result = getDriveItemsByInsights(graph, insightResponse, scopes);
  return result || null;
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
      if (response && response.content) {
        files.push(response.content);
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

async function getDriveItemsByInsights(graph: IGraph, insightResponse, scopes) {
  if (!insightResponse) {
    return [];
  }

  const insightItems = insightResponse.value;
  const batch = graph.createBatch();
  const driveItems = [];
  for (const item of insightItems) {
    const driveItemId = item.resourceReference.id;
    if (driveItemId !== '') {
      batch.get(driveItemId, driveItemId, scopes);
    }
  }

  try {
    const driveItemResponses = await batch.executeAll();

    for (const item of insightItems) {
      const driveItemResponse = driveItemResponses.get(item.resourceReference.id);
      if (driveItemResponse && driveItemResponse.content) {
        driveItems.push(driveItemResponse.content);
      }
    }
    return driveItems;
  } catch (_) {
    try {
      return Promise.all(
        insightItems
          .filter(insightItem => insightItem.resourceReference.id && insightItem.resourceReference.id !== '')
          .map(async insightItem => {
            const driveItemResponses = await graph
              .api(insightItem.resourceReference.id)
              .middlewareOptions(prepScopes(...scopes))
              .get();
            if (driveItemResponses && driveItemResponses.length) {
              return driveItemResponses[0].content;
            }
          })
      );
    } catch (_) {
      return [];
    }
  }
}

async function getFilesPageIterator(graph: IGraph, request) {
  return GraphPageIterator.create<DriveItem>(graph, request);
}
