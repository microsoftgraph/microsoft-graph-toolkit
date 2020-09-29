/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property, LitElement } from 'lit-element';

/**
 * Helper class for SocalizationService
 *
 *
 * @export
 * @class LocalizationHelper
 */
export class LocalizationHelper extends LitElement {
  private static _direction: string;
  public static _strings: object;

  /**
   * returns body dir attribute to determine rtl or ltr
   *
   * @static
   * @returns {string} dir
   * @memberof LocalizationHelper
   */
  public static getDirection() {
    return document.body.getAttribute('dir');
  }

  /**
   * Recieves string and associated tagname from component, compares to new strings
   *
   * @static
   * @param {*} tagName
   * @param {*} stringKey
   * @returns
   * @memberof LocalizationHelper
   */
  public static getString(tagName, stringKey) {
    if (this._strings) {
      let newStringKeys = Object.keys(this._strings);

      if (!this._strings[tagName.toLowerCase()]) {
        return stringKey;
      }

      //converts strings to enum
      function getKeyByValue(object, value) {
        return Object.keys(object).find(key => object[key] === value);
      }

      let enumKey = getKeyByValue(StoredString, stringKey);

      for (let tagKey of newStringKeys) {
        //match tagName
        if (tagKey.toLowerCase() === tagName.toLowerCase()) {
          //checks if tagName has user string
          if (!this._strings[tagName.toLowerCase()][enumKey]) {
            return stringKey;
          }

          //match string
          return this._strings[tagName.toLowerCase()][enumKey];
        }
      }
    }
  }
}

/**
 * StoredString
 *
 * @export
 * @enum {number}
 */
export enum StoredString {
  /**
   * No results found for a picker
   */
  noResultsFound = "We didn't find any matches.",

  /**
   * No results found for a picker
   */
  loading = 'Loading...',

  /**
   * placeholder for mgt-people-picker
   */
  peoplePickerPlaceholder = 'Start typing a name',

  /**
   * Sign in for mgt-login
   */
  signIn = 'Sign In',

  /**
   * Sign out for mgt-login
   */
  signOut = 'Sign Out',

  /**
   * placeholder for mgt-teams-channel-picker
   */
  channelPickerPlaceholder = 'Select a channel',

  /**
   * placeholder for mgt-tasks new task
   */
  newTaskPlaceholder = 'Task...'
}
