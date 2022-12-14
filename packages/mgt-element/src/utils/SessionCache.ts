/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Wrapper around the window.sessionStorage API. Use
 * this to set, get and clear session items.
 */
export class SessionCache {
  private session: Storage;

  constructor() {
    this.session = window.sessionStorage;
  }

  setItem(key: string, value: string) {
    this.session.setItem(key, value);
  }

  getItem(key: string): string {
    return this.session.getItem(key);
  }

  clear() {
    this.session.clear();
  }
}

/**
 * Checks if a sessionStorage or a localStorage is available
 * for use in a browser.
 * @param storageType can be 'sessionStorage' or 'localStorage'.
 * @returns true if the storage is available for use.
 */
export function storageAvailable(storageType: string): boolean {
  let storage: Storage;
  try {
    storage = window[storageType];
    const x = '__storage_test__';
    storage.setItem(x, x);
    storage.removeItem(x);
    return true;
  } catch (e) {
    return (
      e instanceof DOMException &&
      // everything except Firefox
      (e.code === 22 ||
        // Firefox
        e.code === 1014 ||
        // test name field too, because code might not be present
        // everything except Firefox
        e.name === 'QuotaExceededError' ||
        // Firefox
        e.name === 'NS_ERROR_DOM_QUOTA_REACHED') &&
      // acknowledge QuotaExceededError only if there's something already stored
      storage &&
      storage.length !== 0
    );
  }
}
