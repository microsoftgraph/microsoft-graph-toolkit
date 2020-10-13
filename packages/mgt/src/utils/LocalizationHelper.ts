/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher, EventHandler } from '@microsoft/mgt-element';

/**
 * Helper class for Localization
 *
 *
 * @export
 * @class LocalizationHelper
 */
export class LocalizationHelper {
  private static _strings: object;

  private static _eventDispatcher: EventDispatcher<any> = new EventDispatcher();

  public static get strings() {
    return this._strings;
  }

  /**
   * Set strings to be localized
   *
   * @static
   * @memberof LocalizationHelper
   */
  public static set strings(value: any) {
    this._strings = value;
    this._eventDispatcher.fire(null);
  }

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
   * Fires event when LocalizationHelper changes state
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof LocalizationHelper
   */
  public static onUpdated(event: EventHandler<any>) {
    this._eventDispatcher.add(event);
  }

  public static removeOnUpdated(event: EventHandler<any>) {
    this._eventDispatcher.remove(event);
  }

  /**
   * Provided helper method to determine localized or defaultString for specific string is returned
   *
   * @static
   * @param {string} tagName
   * @param  strings
   * @returns
   * @memberof LocalizationHelper
   */
  public static getString(tagName: string, stringObj) {
    tagName = tagName.toLowerCase();

    if (tagName.startsWith('mgt-')) {
      tagName = tagName.substring(4);
    }

    if (this._strings) {
      //check for top level strings, applied per component, overridden by specific component def
      for (let prop of Object.entries(stringObj)) {
        if (this._strings[prop[0]]) {
          stringObj[prop[0]] = this._strings[prop[0]];
        }
      }
      //strings defined component specific
      if (this._strings['_components'] && this._strings['_components'][tagName]) {
        let strings: any = this._strings['_components'][tagName];
        for (let key of Object.keys(strings)) {
          if (stringObj[key]) {
            stringObj[key] = strings[key];
          }
        }
      }
    }

    return stringObj;
  }
}
