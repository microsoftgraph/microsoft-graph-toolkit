/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  CacheItem,
  CacheService,
  CacheStore,
  CollectionResponse,
  GraphPageIterator,
  IGraph,
  prepScopes
} from '@microsoft/mgt-element';
import { DriveItem, SharedInsight, Trending, UploadSession, UsedInsight } from '@microsoft/microsoft-graph-types';
import { schemas } from './cacheStores';
import { GraphRequest, ResponseType } from '@microsoft/microsoft-graph-client';
import { blobToBase64 } from '../utils/Utils';

/**
 * Simple type guard to check if a response is an UploadSession
 *
 * @param session
 * @returns
 */
export const isUploadSession = (session: unknown): session is UploadSession => {
  return Array.isArray((session as UploadSession).nextExpectedRanges);
};

type Insight = SharedInsight | UsedInsight | Trending;

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
export const clearFilesCache = async (): Promise<void> => {
  const cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(
    schemas.fileLists,
    schemas.fileLists.stores.fileLists
  );
  await cache.clearStore();
};

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

/**
 * Load a DriveItem give and arbitrary query
 *
 * @param graph
 * @param resource
 * @returns
 */
export const getDriveItemByQuery = async (
  graph: IGraph,
  resource: string,
  storeName: string = schemas.files.stores.fileQueries,
  scopes = 'files.read'
): Promise<DriveItem> => {
  // get from cache
  const cache: CacheStore<CacheFile> = CacheService.getCache<CacheFile>(schemas.files, storeName);
  const cachedFile = await getFileFromCache(cache, resource);
  if (cachedFile) {
    return cachedFile;
  }

  let response: DriveItem;
  try {
    response = (await graph.api(resource).middlewareOptions(prepScopes(scopes)).get()) as DriveItem;

    if (getIsFilesCacheEnabled()) {
      await cache.putValue(resource, { file: JSON.stringify(response) });
    }
    // eslint-disable-next-line no-empty
  } catch {}

  return response || null;
};

// GET /drives/{drive-id}/items/{item-id}
export const getDriveItemById = async (graph: IGraph, driveId: string, itemId: string): Promise<DriveItem> => {
  const endpoint = `/drives/${driveId}/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.driveFiles);
};

// GET /drives/{drive-id}/root:/{item-path}
export const getDriveItemByPath = async (graph: IGraph, driveId: string, itemPath: string): Promise<DriveItem> => {
  const endpoint = `/drives/${driveId}/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.driveFiles);
};

// GET /groups/{group-id}/drive/items/{item-id}
export const getGroupDriveItemById = async (graph: IGraph, groupId: string, itemId: string): Promise<DriveItem> => {
  const endpoint = `/groups/${groupId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.groupFiles);
};

// GET /groups/{group-id}/drive/root:/{item-path}
export const getGroupDriveItemByPath = async (graph: IGraph, groupId: string, itemPath: string): Promise<DriveItem> => {
  const endpoint = `/groups/${groupId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.groupFiles);
};

// GET /me/drive/items/{item-id}
export const getMyDriveItemById = async (graph: IGraph, itemId: string): Promise<DriveItem> => {
  const endpoint = `/me/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /me/drive/root:/{item-path}
export const getMyDriveItemByPath = async (graph: IGraph, itemPath: string): Promise<DriveItem> => {
  const endpoint = `/me/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /sites/{site-id}/drive/items/{item-id}
export const getSiteDriveItemById = async (graph: IGraph, siteId: string, itemId: string): Promise<DriveItem> => {
  const endpoint = `/sites/${siteId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /sites/{site-id}/drive/root:/{item-path}
export const getSiteDriveItemByPath = async (graph: IGraph, siteId: string, itemPath: string): Promise<DriveItem> => {
  const endpoint = `/sites/${siteId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /sites/{site-id}/lists/{list-id}/items/{item-id}/driveItem
export const getListDriveItemById = async (
  graph: IGraph,
  siteId: string,
  listId: string,
  itemId: string
): Promise<DriveItem> => {
  const endpoint = `/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.siteFiles);
};

// GET /users/{user-id}/drive/items/{item-id}
export const getUserDriveItemById = async (graph: IGraph, userId: string, itemId: string): Promise<DriveItem> => {
  const endpoint = `/users/${userId}/drive/items/${itemId}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /users/{user-id}/drive/root:/{item-path}
export const getUserDriveItemByPath = async (graph: IGraph, userId: string, itemPath: string): Promise<DriveItem> => {
  const endpoint = `/users/${userId}/drive/root:/${itemPath}`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.userFiles);
};

// GET /me/insights/trending/{id}/resource
// GET /me/insights/used/{id}/resource
// GET /me/insights/shared/{id}/resource
export const getMyInsightsDriveItemById = async (
  graph: IGraph,
  insightType: string,
  id: string
): Promise<DriveItem> => {
  const endpoint = `/me/insights/${insightType}/${id}/resource`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.insightFiles, 'sites.read.all');
};

// GET /users/{id or userPrincipalName}/insights/{trending or used or shared}/{id}/resource
export const getUserInsightsDriveItemById = async (
  graph: IGraph,
  userId: string,
  insightType: string,
  id: string
): Promise<DriveItem> => {
  const endpoint = `/users/${userId}/insights/${insightType}/${id}/resource`;
  return getDriveItemByQuery(graph, endpoint, schemas.files.stores.insightFiles, 'sites.read.all');
};

const getIterator = async (
  graph: IGraph,
  endpoint: string,
  storeName: string,
  scopes: string[],
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  let filesPageIterator: GraphPageIterator<DriveItem>;

  // get iterator from cached values
  const cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, storeName);
  const fileList = await getFileListFromCache(cache, storeName, `${endpoint}:${top}`);
  if (fileList) {
    filesPageIterator = getFilesPageIteratorFromCache(graph, fileList.files, fileList.nextLink);

    return filesPageIterator;
  }

  // get iterator from graph request
  let request: GraphRequest;
  try {
    request = graph.api(endpoint).middlewareOptions(prepScopes(...scopes));
    if (top) {
      request.top(top);
    }
    filesPageIterator = await getFilesPageIteratorFromRequest(graph, request);

    if (getIsFileListsCacheEnabled()) {
      const nextLink = filesPageIterator.nextLink;
      await cache.putValue(endpoint, {
        files: filesPageIterator.value.map(v => JSON.stringify(v)),
        nextLink
      });
    }
    // eslint-disable-next-line no-empty
  } catch {}
  return filesPageIterator || null;
};

// GET /me/drive/root/children
export const getFilesIterator = async (graph: IGraph, top?: number): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = '/me/drive/root/children';
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /drives/{drive-id}/items/{item-id}/children
export const getDriveFilesByIdIterator = async (
  graph: IGraph,
  driveId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/drives/${driveId}/items/${itemId}/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /drives/{drive-id}/root:/{item-path}:/children
export const getDriveFilesByPathIterator = async (
  graph: IGraph,
  driveId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/drives/${driveId}/root:/${itemPath}:/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /groups/{group-id}/drive/items/{item-id}/children
export const getGroupFilesByIdIterator = async (
  graph: IGraph,
  groupId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/groups/${groupId}/drive/items/${itemId}/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /groups/{group-id}/drive/root:/{item-path}:/children
export const getGroupFilesByPathIterator = async (
  graph: IGraph,
  groupId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/groups/${groupId}/drive/root:/${itemPath}:/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /me/drive/items/{item-id}/children
export const getFilesByIdIterator = async (
  graph: IGraph,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/me/drive/items/${itemId}/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /me/drive/root:/{item-path}:/children
export const getFilesByPathIterator = async (
  graph: IGraph,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/me/drive/root:/${itemPath}:/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /sites/{site-id}/drive/items/{item-id}/children
export const getSiteFilesByIdIterator = async (
  graph: IGraph,
  siteId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/sites/${siteId}/drive/items/${itemId}/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /sites/{site-id}/drive/root:/{item-path}:/children
export const getSiteFilesByPathIterator = async (
  graph: IGraph,
  siteId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/sites/${siteId}/drive/root:/${itemPath}:/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /users/{user-id}/drive/items/{item-id}/children
export const getUserFilesByIdIterator = async (
  graph: IGraph,
  userId: string,
  itemId: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/users/${userId}/drive/items/${itemId}/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

// GET /users/{user-id}/drive/root:/{item-path}:/children
export const getUserFilesByPathIterator = async (
  graph: IGraph,
  userId: string,
  itemPath: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const endpoint = `/users/${userId}/drive/root:/${itemPath}:/children`;
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, endpoint, cacheStore, ['files.read'], top);
};

export const getFilesByListQueryIterator = async (
  graph: IGraph,
  listQuery: string,
  top?: number
): Promise<GraphPageIterator<DriveItem>> => {
  const cacheStore = schemas.fileLists.stores.fileLists;
  return getIterator(graph, listQuery, cacheStore, ['files.read', 'sites.read.all'], top);
};

// GET /me/insights/{trending	| used | shared}
export const getMyInsightsFiles = async (graph: IGraph, insightType: string): Promise<DriveItem[]> => {
  const endpoint = `/me/insights/${insightType}`;
  const cacheStore = schemas.fileLists.stores.insightfileLists;

  // get files from cached values
  const cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, endpoint);
  if (fileList) {
    // fileList.files is string[] so JSON.parse to get proper objects
    return fileList.files.map((file: string) => JSON.parse(file) as DriveItem);
  }

  // get files from graph request
  const scopes = ['sites.read.all'];
  let insightResponse: CollectionResponse<Insight>;
  try {
    insightResponse = (await graph
      .api(endpoint)
      .filter("resourceReference/type eq 'microsoft.graph.driveItem'")
      .middlewareOptions(prepScopes(...scopes))
      .get()) as CollectionResponse<Insight>;
    // eslint-disable-next-line no-empty
  } catch {}

  const result = await getDriveItemsByInsights(graph, insightResponse, scopes);
  if (getIsFileListsCacheEnabled()) {
    await cache.putValue(endpoint, { files: result.map(file => JSON.stringify(file)) });
  }

  return result || null;
};

// GET /users/{id | userPrincipalName}/insights/{trending	| used | shared}
export const getUserInsightsFiles = async (
  graph: IGraph,
  userId: string,
  insightType: string
): Promise<DriveItem[]> => {
  let endpoint: string;
  let filter: string;

  if (insightType === 'shared') {
    endpoint = '/me/insights/shared';
    filter = `((lastshared/sharedby/id eq '${userId}') and (resourceReference/type eq 'microsoft.graph.driveItem'))`;
  } else {
    endpoint = `/users/${userId}/insights/${insightType}`;
    filter = "resourceReference/type eq 'microsoft.graph.driveItem'";
  }

  const key = `${endpoint}?$filter=${filter}`;

  // get files from cached values
  const cacheStore = schemas.fileLists.stores.insightfileLists;
  const cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(schemas.fileLists, cacheStore);
  const fileList = await getFileListFromCache(cache, cacheStore, key);
  if (fileList) {
    return fileList.files.map((file: string) => JSON.parse(file) as DriveItem);
  }

  // get files from graph request
  const scopes = ['sites.read.all'];
  let insightResponse: CollectionResponse<Insight>;

  try {
    insightResponse = (await graph
      .api(endpoint)
      .filter(filter)
      .middlewareOptions(prepScopes(...scopes))
      .get()) as CollectionResponse<Insight>;
    // eslint-disable-next-line no-empty
  } catch {}

  const result = await getDriveItemsByInsights(graph, insightResponse, scopes);
  if (getIsFileListsCacheEnabled()) {
    await cache.putValue(endpoint, { files: result.map(file => JSON.stringify(file)) });
  }

  return result || null;
};

export const getFilesByQueries = async (graph: IGraph, fileQueries: string[]): Promise<DriveItem[]> => {
  if (!fileQueries || fileQueries.length === 0) {
    return [];
  }

  const batch = graph.createBatch();
  const files: DriveItem[] = [];
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
      files.push(JSON.parse(cachedFile.file) as DriveItem);
    } else if (fileQuery !== '') {
      batch.get(fileQuery, fileQuery, scopes);
    }
  }

  try {
    const responses = await batch.executeAll();

    for (const fileQuery of fileQueries) {
      const response = responses.get(fileQuery);
      if (response?.content) {
        files.push(response.content as DriveItem);
        if (getIsFilesCacheEnabled()) {
          await cache.putValue(fileQuery, { file: JSON.stringify(response.content) });
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
                await cache.putValue(fileQuery, { file: JSON.stringify(file) });
              }
              return file;
            }
          })
      );
    } catch (e) {
      return [];
    }
  }
};

const getDriveItemsByInsights = async (
  graph: IGraph,
  insightResponse: CollectionResponse<Insight>,
  scopes: string[]
): Promise<DriveItem[]> => {
  if (!insightResponse) {
    return [];
  }

  const insightItems = insightResponse.value;
  const batch = graph.createBatch();
  const driveItems: DriveItem[] = [];
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
      if (driveItemResponse?.content) {
        driveItems.push(driveItemResponse.content as DriveItem);
      }
    }
    return driveItems;
  } catch (_) {
    try {
      // we're filtering the insights calls that feed this to ensure that only
      // drive items are returned, but we still need to check for nulls
      return Promise.all(
        insightItems
          .filter(insightItem => Boolean(insightItem.resourceReference.id))
          .map(
            async insightItem =>
              (await graph
                .api(insightItem.resourceReference.id)
                .middlewareOptions(prepScopes(...scopes))
                .get()) as DriveItem
          )
      );
    } catch (e) {
      return [];
    }
  }
};

const getFilesPageIteratorFromRequest = async (graph: IGraph, request: GraphRequest) => {
  return GraphPageIterator.create<DriveItem>(graph, request);
};

const getFilesPageIteratorFromCache = (graph: IGraph, value: string[], nextLink: string) => {
  return GraphPageIterator.createFromValue<DriveItem>(
    graph,
    value.map(v => JSON.parse(v) as DriveItem),
    nextLink
  );
};

/**
 * Load a file from the cache
 *
 * @param {CacheStore<CacheFile>} cache
 * @param {string} key
 * @return {*}
 */
const getFileFromCache = async <TResult = DriveItem>(cache: CacheStore<CacheFile>, key: string): Promise<TResult> => {
  if (getIsFilesCacheEnabled()) {
    const file = await cache.getValue(key);

    if (file && getFileInvalidationTime() > Date.now() - file.timeCached) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-return
      return JSON.parse(file.file) as TResult;
    }
  }

  return null;
};

export const getFileListFromCache = async (cache: CacheStore<CacheFileList>, store: string, key: string) => {
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
};

// refresh filesPageIterator to its next iteration and save current page to cache
export const fetchNextAndCacheForFilesPageIterator = async (filesPageIterator: GraphPageIterator<DriveItem>) => {
  const nextLink = filesPageIterator.nextLink;

  if (filesPageIterator.hasNext) {
    await filesPageIterator.next();
  }
  if (getIsFileListsCacheEnabled()) {
    const cache: CacheStore<CacheFileList> = CacheService.getCache<CacheFileList>(
      schemas.fileLists,
      schemas.fileLists.stores.fileLists
    );

    // match only the endpoint (after version number and before OData query params) e.g. /me/drive/root/children
    const reg = /(graph.microsoft.com\/(v1.0|beta))(.*?)(?=\?)/gi;
    const matches = reg.exec(nextLink);
    const key = matches[3];

    await cache.putValue(key, { files: filesPageIterator.value.map(v => JSON.stringify(v)), nextLink });
  }
};

/**
 * retrieves the specified document thumbnail
 *
 * @param {string} resource
 * @param {string[]} scopes
 * @returns {Promise<string>}
 */
export const getDocumentThumbnail = async (
  graph: IGraph,
  resource: string,
  scopes: string[]
): Promise<CacheThumbnail> => {
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
};

/**
 * retrieve file properties based on Graph query
 *
 * @param graph
 * @param resource
 * @returns
 */
export const getGraphfile = async (graph: IGraph, resource: string): Promise<DriveItem> => {
  // get from graph request
  const scopes = 'files.read';
  try {
    const response = (await graph.api(resource).middlewareOptions(prepScopes(scopes)).get()) as DriveItem;
    return response || null;
    // eslint-disable-next-line no-empty
  } catch {}

  return null;
};

/**
 * retrieve UploadSession Url for large file and send by chuncks
 *
 * @param graph
 * @param resource
 * @returns
 */
export const getUploadSession = async (
  graph: IGraph,
  resource: string,
  conflictBehavior: number
): Promise<UploadSession> => {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    const sessionOptions = {
      item: {
        '@microsoft.graph.conflictBehavior': conflictBehavior === 0 || conflictBehavior === null ? 'rename' : 'replace'
      }
    };
    let response: UploadSession;
    try {
      response = (await graph
        .api(resource)
        .middlewareOptions(prepScopes(scopes))
        .post(JSON.stringify(sessionOptions))) as UploadSession;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * send file chunck to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export const sendFileChunk = async (
  graph: IGraph,
  resource: string,
  contentLength: string,
  contentRange: string,
  file: Blob
): Promise<UploadSession | DriveItem> => {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    const header = {
      'Content-Length': contentLength,
      'Content-Range': contentRange
    };
    let response: UploadSession | DriveItem;
    try {
      response = (await graph.client.api(resource).middlewareOptions(prepScopes(scopes)).headers(header).put(file)) as
        | UploadSession
        | DriveItem;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * send file to OneDrive, SharePoint Site
 *
 * @param graph
 * @param resource
 * @param file
 * @returns
 */
export const sendFileContent = async (graph: IGraph, resource: string, file: File): Promise<DriveItem> => {
  try {
    // get from graph request
    const scopes = 'files.readwrite';
    let response: DriveItem;
    try {
      response = (await graph.client.api(resource).middlewareOptions(prepScopes(scopes)).put(file)) as DriveItem;
      // eslint-disable-next-line no-empty
    } catch {}

    return response || null;
  } catch (e) {
    return null;
  }
};

/**
 * delete upload session
 *
 * @param graph
 * @param resource
 * @returns
 */
export const deleteSessionFile = async (graph: IGraph, resource: string): Promise<void> => {
  const scopes = 'files.readwrite';
  try {
    await graph.client.api(resource).middlewareOptions(prepScopes(scopes)).delete();
  } catch {
    // TODO: re-examine the error handling here
    // DELETE returns a 204 on success so void makes sense to return on the happy path
    // but we should probably throw on error
    return null;
  }
};
