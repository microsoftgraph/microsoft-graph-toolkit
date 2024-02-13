/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

export const replaceOrAppendSessionId = (url: string, sessionId: string) => {
  // match on any whole query parameter of sessionid and include the value
  const sessionIdRegex = /([?&]{1})sessionid=[^&]*/;
  if (url.match(sessionIdRegex)) {
    url = url.replace(sessionIdRegex, `$1sessionid=${sessionId}`);
  } else {
    const paramSeparator = url.indexOf('?') > -1 ? '&' : '?';
    url = `${url}${paramSeparator}sessionid=${sessionId}`;
  }
  return url;
};
