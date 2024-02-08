/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { v4 as uuid } from 'uuid';

const keyPrefix = 'mgt-chat-group-id';

/**
 * reads a string from session storage, or if there is no string for the keyName, generate a new uuid and place in storage
 */
export const getOrGenerateGroupId = (chatId: string) => {
  const key = `${keyPrefix}::${chatId}`;
  const value = localStorage.getItem(key);

  if (value) {
    return value;
  } else {
    const newValue = uuid();
    localStorage.setItem(key, newValue);
    return newValue;
  }
};
