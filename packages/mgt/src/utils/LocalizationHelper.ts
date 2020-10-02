/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher, EventHandler } from '@microsoft/mgt-element';
import * as defaultStrings from '../l10n/en-us';

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
   * Fires event when Provider changes state
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof Providers
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
   * @param {string} stringKey
   * @returns
   * @memberof LocalizationHelper
   */
  public static getString(tagName: string, stringKey: string) {
    let string: string;

    tagName = tagName.toLowerCase();

    if (tagName.startsWith('mgt-')) {
      tagName = tagName.substring(4);
    }

    // first search through the user provided strings
    if (this._strings) {
      if (this._strings['_components'] && this._strings['_components'][tagName]) {
        let strings = this._strings['_components'][tagName];
        string = strings[stringKey];
      }

      if (!string) {
        string = this._strings[stringKey];
      }
    }

    // if no stringKey in user provider strings, or no user provider strings
    if (!string) {
      if (defaultStrings.default['_components'] && defaultStrings.default['_components'][tagName]) {
        let strings = defaultStrings.default['_components'][tagName];
        string = strings[stringKey];
      }

      if (!string) {
        string = defaultStrings.default[stringKey];
      }
    }

    return string;
  }
}
