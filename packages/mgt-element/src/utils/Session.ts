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
export class Session {
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
