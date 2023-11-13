/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

export const updateMessageContentWithImage = (
  content: string,
  index: string,
  messageId: string,
  image: string
): string => {
  const imageId = `${messageId}-${index}`;
  const regex = new RegExp(`(<img[^>]+id=["']${imageId}["'][^>]+)src=["']([^"']*)["']`);
  const replacement = `$1src="${image}"`;
  const replaced = content.replace(regex, replacement);
  return replaced;
};
