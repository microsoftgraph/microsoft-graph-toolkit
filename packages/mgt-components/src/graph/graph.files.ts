/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';

export async function getDriveItemByQuery(graph: IGraph, resource: string): Promise<DriveItem> {
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Files.Read';
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
  const scopes = 'Sites.Read.All';
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
  const scopes = 'Sites.Read.All';
  let response;
  try {
    response = await graph
      .api(`/users/${userId}/insights/${insightType}/${id}/resource`)
      .middlewareOptions(prepScopes(scopes))
      .get();
  } catch {}
  return response || null;
}
