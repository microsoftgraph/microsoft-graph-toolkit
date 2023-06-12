/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const fileTypeMap: Record<string, string> = {
  PowerPoint: 'pptx',
  Word: 'docx',
  Excel: 'xlsx',
  Pdf: 'pdf',
  OneNote: 'onetoc',
  OneNotePage: 'one',
  InfoPath: 'xsn',
  Visio: 'vstx',
  Publisher: 'pub',
  Project: 'mpp',
  Access: 'accdb',
  Mail: 'email',
  Csv: 'csv',
  Archive: 'archive',
  Xps: 'vector',
  Audio: 'audio',
  Video: 'video',
  Image: 'photo',
  Web: 'html',
  Text: 'txt',
  Xml: 'xml',
  Story: 'genericfile',
  ExternalContent: 'genericfile',
  Folder: 'folder',
  Spsite: 'spo',
  Other: 'genericfile'
};

const baseUri = 'https://spoprod-a.akamaihd.net/files/fabric-cdn-prod_20201008.001/assets/item-types';

/**
 * Available icon sizes
 */
export type IconSize = 16 | 20 | 24 | 32 | 40 | 48 | 64 | 96;

/**
 * Helper to provide fluent icon image urls
 *
 * @param type
 * @param size
 * @param extension
 * @returns
 */
export const getFileTypeIconUri = (type: string, size: IconSize, extension: 'png' | 'svg') => {
  const fileType: string = fileTypeMap[type] || 'genericfile';
  return `${baseUri}/${size.toString()}/${fileType}.${extension}`;
};

/**
 * Helper to provide fluent icon image urls with the correct size
 *
 * @param type
 * @param size
 * @param extension
 * @returns
 */
export const getFileTypeIconUriByExtension = (type: string, size: IconSize, extension: 'png' | 'svg') => {
  const found = Object.keys(fileTypeMap).find(key => fileTypeMap[key] === type);
  if (found) {
    return `${baseUri}/${size.toString()}/${type}.${extension}`;
  } else if (type === 'jpg' || type === 'png') {
    type = 'photo';
    return `${baseUri}/${size.toString()}/${type}.${extension}`;
  } else {
    return null;
  }
};
