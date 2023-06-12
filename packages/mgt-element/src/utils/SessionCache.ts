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
  private readonly session: Storage;

  constructor() {
    this.session = window.sessionStorage;
  }

  /**
   * Stores a value in session storage.
   *
   * @param key
   * @param value
   */
  setItem(key: string, value: string) {
    this.session.setItem(key, value);
  }

  /**
   * Gets the value for a given key from session storage.
   *
   * @param {string} key
   * @return {*}  {string}
   * @memberof SessionCache
   */
  getItem(key: string): string {
    return this.session.getItem(key);
  }

  /**
   * Clears session storage.
   *
   * @memberof SessionCache
   */
  clear() {
    this.session.clear();
  }
}

/**
 * Checks if a sessionStorage or a localStorage is available
 * for use in a browser.
 *
 * @param storageType can be 'sessionStorage' or 'localStorage'.
 * @returns true if the storage is available for use.
 */
export const storageAvailable = (storageType: string): boolean => {
  let storage: Storage;
  try {
    if (storageType === 'sessionStorage') {
      storage = window.sessionStorage;
    } else {
      storage = window.localStorage;
    }
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
};
