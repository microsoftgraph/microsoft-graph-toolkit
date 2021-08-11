/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheItem, CacheService, CacheStore, GraphPageIterator, IGraph, prepScopes } from '@microsoft/mgt-element';
import { DriveItem, UploadSession } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { blobToBase64 } from '../utils/Utils';

/**
 * Object to be stored in cache
 */
interface CacheFile extends CacheItem {
  /**
   * stringified json representing a file
   */
  file?: string;
}

/**
 * Object to be stored in cache
 */
interface CacheFileList extends CacheItem {
  /**
   * stringified json representing a list of files
   */
  files?: string[];
  /**
   * nextLink string to get next page
   */
  nextLink?: string;
}

/**
 * document thumbnail object stored in cache
 */
export interface CacheThumbnail extends CacheItem {
  /**
   * tag associated with thumbnail
   */
  eTag?: string;
  /**
   * document thumbnail
   */
  thumbnail?: string;
}

/**
 * Clear Cache of FileList
 */
export function clearFilesCache() {
  let cache: CacheStore<CacheFileList>;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, schemas.fileLists.stores.fileLists);
  cache.clearStore();
}

/**
 * Defines the time it takes for objects in the cache to expire
 */
export const getFileInvalidationTime = (): number =>
  CacheService.config.files.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
export const getIsFilesCacheEnabled = (): boolean =>
  CacheService.config.files.isEnabled && CacheService.config.isEnabled;

/**
 * Defines the time it takes for objects in the cache to expire
 */
export const getFileListInvalidationTime = (): number =>
  CacheService.config.fileLists.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * Whether or not the cache is enabled
 */
export const getIsFileListsCacheEnabled = (): boolean =>
  CacheService.config.fileLists.isEnabled && CacheService.config.isEnabled;

export async function getDriveItemByQuery(graph: IGraph, resource: string): Promise<DriveItem> {
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.fileQueries);
  const cachedFile = await getFileFromCache(cache, resource);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(resource).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(resource, { file: JSON.stringify(response) });
    }
  } catch {}

  return response || null;
}

// GET /drives/{drive-id}/items/{item-id}
export async function getDriveItemById(graph: IGraph, driveId: string, itemId: string): Promise<DriveItem> {
  const endpoint = `/drives/${driveId}/items/${itemId}`;

  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.driveFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /drives/{drive-id}/root:/{item-path}
export async function getDriveItemByPath(graph: IGraph, driveId: string, itemPath: string): Promise<DriveItem> {
  const endpoint = `/drives/${driveId}/root:/${itemPath}`;

  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.driveFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /groups/{group-id}/drive/items/{item-id}
export async function getGroupDriveItemById(graph: IGraph, groupId: string, itemId: string): Promise<DriveItem> {
  const endpoint = `/groups/${groupId}/drive/items/${itemId}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.groupFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /groups/{group-id}/drive/root:/{item-path}
export async function getGroupDriveItemByPath(graph: IGraph, groupId: string, itemPath: string): Promise<DriveItem> {
  const endpoint = `/groups/${groupId}/drive/root:/${itemPath}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.groupFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /me/drive/items/{item-id}
export async function getMyDriveItemById(graph: IGraph, itemId: string): Promise<DriveItem> {
  const endpoint = `/me/drive/items/${itemId}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.userFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /me/drive/root:/{item-path}
export async function getMyDriveItemByPath(graph: IGraph, itemPath: string): Promise<DriveItem> {
  const endpoint = `/me/drive/root:/${itemPath}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.userFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /sites/{site-id}/drive/items/{item-id}
export async function getSiteDriveItemById(graph: IGraph, siteId: string, itemId: string): Promise<DriveItem> {
  const endpoint = `/sites/${siteId}/drive/items/${itemId}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.siteFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /sites/{site-id}/drive/root:/{item-path}
export async function getSiteDriveItemByPath(graph: IGraph, siteId: string, itemPath: string): Promise<DriveItem> {
  const endpoint = `/sites/${siteId}/drive/root:/${itemPath}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.siteFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
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
  const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.siteFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /users/{user-id}/drive/items/{item-id}
export async function getUserDriveItemById(graph: IGraph, userId: string, itemId: string): Promise<DriveItem> {
  const endpoint = `/users/${userId}/drive/items/${itemId}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.userFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /users/{user-id}/drive/root:/{item-path}
export async function getUserDriveItemByPath(graph: IGraph, userId: string, itemPath: string): Promise<DriveItem> {
  const endpoint = `/users/${userId}/drive/root:/${itemPath}`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.userFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'files.read';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /me/insights/trending/{id}/resource
// GET /me/insights/used/{id}/resource
// GET /me/insights/shared/{id}/resource
export async function getMyInsightsDriveItemById(graph: IGraph, insightType: string, id: string): Promise<DriveItem> {
  const endpoint = `/me/insights/${insightType}/${id}/resource`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.insightFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
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
  const endpoint = `/users/${userId}/insights/${insightType}/${id}/resource`;
  // get from cache
  let cache: CacheStore<CacheFile>;
  cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.insightFiles);
  const cachedFile = await getFileFromCache(cache, endpoint);
  if (cachedFile) {
    return cachedFile;
  }

  // get from graph request
  const scopes = 'sites.read.all';
  let response;
  try {
    response = await graph.api(endpoint).middlewareOptions(prepScopes(scopes)).get();

    if (getIsFilesCacheEnabled()) {
      cache.putValue(endpoint, { file: JSON.stringify(response) });
    }
  } catch {}
  return response || null;
}

// GET /me/drive/root/children
export async function getFilesIterator(graph: IGraph, top?: number): Promise<GraphPageIterator<DriveItem>> {
  const endpoint = '/me/drive/root/children';
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/drives/${driveId}/items/${itemId}/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/drives/${driveId}/root:/${itemPath}:/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/groups/${groupId}/drive/items/${itemId}/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/groups/${groupId}/drive/root:/${itemPath}:/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
  } catch {}
  return filesPageIterator || null;
}

// GET /me/drive/items/{item-id}/children
export async function getFilesByIdIterator(
  graph: IGraph,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const endpoint = `/me/drive/items/${itemId}/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
  } catch {}
  return filesPageIterator || null;
}

// GET /me/drive/root:/{item-path}:/children
export async function getFilesByPathIterator(
  graph: IGraph,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  const endpoint = `/me/drive/root:/${itemPath}:/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/sites/${siteId}/drive/items/${itemId}/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/sites/${siteId}/drive/root:/${itemPath}:/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/users/${userId}/drive/items/${itemId}/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
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
  const endpoint = `/users/${userId}/drive/root:/${itemPath}:/children`;
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = 'files.read';
  let request;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(endpoint, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
  } catch {}
  return filesPageIterator || null;
}

export async function getFilesByListQueryIterator(
  graph: IGraph,
  listQuery: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> {
  let filesPageIterator;

  // get iterator from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.fileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, listQuery);
  if (fileList) {
    filesPageIterator = await getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  const scopes = ['files.read', 'sites.read.all'];
  let request;

  try {
    request = await graph.api(listQuery).middlewareOptions(prepScopes(...scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      cache.putValue(listQuery, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
    }
  } catch {}
  return filesPageIterator || null;
}

// GET /me/insights/{trending	| used | shared}
export async function getMyInsightsFiles(graph: IGraph, insightType: string): Promise<DriveItem[]> {
  const endpoint = `/me/insights/${insightType}`;

  // get files from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.insightfileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    return fileList.files as DriveItem[];
  }

  // get files from graph request
  const scopes = ['sites.read.all'];
  let insightResponse;
  try {
    insightResponse = await graph
      .api(endpoint)
      .filter(`resourceReference/type eq 'microsoft.graph.driveItem'`)
      .middlewareOptions(prepScopes(...scopes))
      .get();
  } catch {}

  const result = await getDriveItemsByInsights(graph, insightResponse, scopes);
  if (getIsFileListsCacheEnabled()) {
    cache.putValue(endpoint, { files: result });
  }

  return result || null;
}

// GET /users/{id | userPrincipalName}/insights/{trending	| used | shared}
export async function getUserInsightsFiles(graph: IGraph, userId: string, insightType: string): Promise<DriveItem[]> {
  let endpoint;
  let filter;

  if (insightType === 'shared') {
    endpoint = `/me/insights/shared`;
    filter = `((lastshared/sharedby/id eq '${userId}') and (resourceReference/type eq 'microsoft.graph.driveItem'))`;
  } else {
    endpoint = `/users/${userId}/insights/${insightType}`;
    filter = `resourceReference/type eq 'microsoft.graph.driveItem'`;
  }

  const key = `${endpoint}?$filter=${filter}`;

  // get files from cached values
  let cache: CacheStore<CacheFileList>;
  const cacheStore = schemas.fileLists.stores.insightfileLists;
  cache = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, key);
  if (fileList) {
    return fileList.files as DriveItem[];
  }

  // get files from graph request
  const scopes = ['sites.read.all'];
  let insightResponse;

  try {
    insightResponse = await graph
      .api(endpoint)
      .filter(filter)
      .middlewareOptions(prepScopes(...scopes))
      .get();
  } catch {}

  const result = await getDriveItemsByInsights(graph, insightResponse, scopes);
  if (getIsFileListsCacheEnabled()) {
    cache.putValue(endpoint, { files: result });
  }

  return result || null;
}

export async function getFilesByQueries(graph: IGraph, fileQueries: string[]): Promise<DriveItem[]> {
  if (!fileQueries || fileQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch();
  const files = [];
  const scopes = ['files.read'];
  let cache: CacheStore<CacheFile>;
  let cachedFile: CacheFile;
  if (getIsFilesCacheEnabled()) {
    cache = CacheService.getCache<CacheFile>(schemas.files, schemas.files.stores.fileQueries);
  }

  for (const fileQuery of fileQueries) {
    if (getIsFilesCacheEnabled()) {
      cachedFile = await cache.getValue(fileQuery); // todo
    }

    if (getIsFilesCacheEnabled() && cachedFile && getFileInvalidationTime() > Date.now() - cachedFile.timeCached) {
      files.push(JSON.parse(cachedFile.file));
    } else if (fileQuery !== '') {
      batch.get(fileQuery, fileQuery, scopes);
    }
  }

  try {
    const responses = await batch.executeAll();

    for (const fileQuery of fileQueries) {
      const response = responses.get(fileQuery);
      if (response && response.content) {
        files.push(response.content);
        if (getIsFilesCacheEnabled()) {
          cache.putValue(fileQuery, { file: JSON.stringify(response.content) });
        }
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
              if (getIsFilesCacheEnabled()) {
                cache.putValue(fileQuery, { file: JSON.stringify(file) });
              }
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

async function getFilesPageIteratorFromRequest(graph: IGraph, request) {
  return GraphPageIterator.create<DriveItem>(graph, request);
}

async function getFilesPageIteratorFromCache(graph: IGraph, value, nextLink) {
  return GraphPageIterator.createFromValue<DriveItem>(graph, value, nextLink);
}

async function getFileFromCache(cache: CacheStore<CacheFile>, key: string) {
  if (getIsFilesCacheEnabled()) {
    const file = await cache.getValue(key);

    if (file && getFileInvalidationTime() > Date.now() - file.timeCached) {
      const cachedFile = JSON.parse(file.file);
      return cachedFile;
    }
  }

  return null;
}

export async function getFileListFromCache(cache: CacheStore<CacheFileList>, store: string, key: string) {
  if (!cache) {
    cache = CacheService.getCache<CacheFileList>(schemas.fileLists, store);
  }

  if (getIsFileListsCacheEnabled()) {
    const fileList = await cache.getValue(key);

    if (fileList && getFileListInvalidationTime() > Date.now() - fileList.timeCached) {
      return fileList;
    }
  }

  return null;
}

// refresh filesPageIterator to its next iteration and save current page to cache
export async function fetchNextAndCacheForFilesPageIterator(filesPageIterator) {
  const nextLink = filesPageIterator._nextLink;

  if (filesPageIterator.hasNext) {
    await filesPageIterator.next();
  }
  if (getIsFileListsCacheEnabled()) {
    let cache: CacheStore<CacheFileList>;
    cache = CacheService.getCache<CacheFileList>(schemas.fileLists, schemas.fileLists.stores.fileLists);

    // match only the endpoint (after version number and before OData query params) e.g. /me/drive/root/children
    const reg = /(graph.microsoft.com\/(v1.0|beta))(.*?)(?=\?)/gi;
    const matches = reg.exec(nextLink);
    const key = matches[3];

    cache.putValue(key, { files: filesPageIterator.value, nextLink: filesPageIterator._nextLink });
  }
}

/**
 * retrieves the specified document thumbnail
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
export async function getDocumentThumbnail(graph: IGraph, resource: string, scopes: string[]): Promise<CacheThumbnail> {
  try {
    const response = (await graph
      .api(resource)
      .responseType(ResponseType.RAW)
      .middlewareOptions(prepScopes(...scopes))
      .get()) as Response;

    if (response.status === 404) {
      // 404 means the resource does not have a thumbnail
      // we still want to cache that state
      // so we return an object that can be cached
      return { eTag: null, thumbnail: null };
    } else if (!response.ok) {
      return null;
    }

    const eTag = response.headers.get('eTag');
    const blob = await blobToBase64(await response.blob());
    return { eTag, thumbnail: blob };
  } catch (e) {
    return null;
  }
}

/**
 * retrieve file properties based on Graph query
 *
 * @param graph
 * @param resource
 * @returns
 */
export async function getGraphfile(graph: IGraph, resource: string): Promise<DriveItem> {
  try {
    // get from graph request
    const scopes = 'files.read';
    let response;
    try {
      response = await graph
        .api(resource)
        .middlewareOptions(prepScopes(scopes))
        .get()
        .catch(error => {
          return null;
        });
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
}

/**
 * retrieve UploadSession Url for large file and send by chuncks
 *
 * @param graph
 * @param resource
 * @returns
 */
export async function getUploadSession(
  graph: IGraph,
  resource: string,
  conflictBehavior: number
): Promise<UploadSession> {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    const sessionOptions = {
      item: {
        '@microsoft.graph.conflictBehavior': conflictBehavior === 0 || conflictBehavior === null ? 'rename' : 'replace'
      }
    };
    let response;
    try {
      response = await graph.api(resource).middlewareOptions(prepScopes(scopes)).post(JSON.stringify(sessionOptions));
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
}

/**
 * send file chunck to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export async function sendFileChunck(
  graph: IGraph,
  resource: string,
  contentLength: string,
  contentRange: string,
  file: Blob
): Promise<any> {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    const header = {
      'Content-Length': contentLength,
      'Content-Range': contentRange
    };
    let response;
    try {
      response = await graph.client.api(resource).middlewareOptions(prepScopes(scopes)).headers(header).put(file);
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
}

/**
 * send file to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export async function sendFileContent(graph: IGraph, resource: string, file: File): Promise<DriveItem> {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    let response;
    try {
      response = await graph.client.api(resource).middlewareOptions(prepScopes(scopes)).put(file);
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
}

/**
 * delete upload session
 *
 * @param graph
 * @param resource
 * @returns
 */
export async function deleteSessionFile(graph: IGraph, resource: string): Promise<any> {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    let response;
    try {
      response = await graph.client
        .api(resource)
        .middlewareOptions(prepScopes(scopes))
        .delete(response => {
          return response;
        });
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
}
