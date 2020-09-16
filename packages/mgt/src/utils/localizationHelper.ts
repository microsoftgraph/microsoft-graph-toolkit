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

  public static getString(tagName, stringKey) {
    console.log(tagName, stringKey);
    console.log(this._strings);
    if (this._strings) {
      let newStringKeys = Object.keys(this._strings);

      if (!this._strings[tagName.toLowerCase()]) {
        return stringKey;
      }

      for (let tagKey of newStringKeys) {
        //match tagName
        if (tagKey.toLowerCase() === tagName.toLowerCase()) {
          //match string
          return this._strings[tagName.toLowerCase()][stringKey];
        }
      }
    }
  }
}
