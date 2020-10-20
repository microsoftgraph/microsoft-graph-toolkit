/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher, EventHandler } from './EventDispatcher';

/**
 * Helper class for Localization
 *
 *
 * @export
 * @class LocalizationHelper
 */
export class LocalizationHelper {
  static _strings: any;

  static _stringsEventDispatcher: EventDispatcher<any> = new EventDispatcher();

  static _directionEventDispatcher: EventDispatcher<any> = new EventDispatcher();

  private static mutationObserver;

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
    this._stringsEventDispatcher.fire(null);
  }

  /**
   * returns body dir attribute to determine rtl or ltr
   *
   * @static
   * @returns {string} dir
   * @memberof LocalizationHelper
   */
  public static getDocumentDirection() {
    this.initDirection();
    return document.body.getAttribute('dir') || document.documentElement.getAttribute('dir');
  }

  /**
   * Fires event when LocalizationHelper changes state
   *
   * @static
   * @param {EventHandler<ProvidersChangedState>} event
   * @memberof LocalizationHelper
   */
  public static onStringsUpdated(event: EventHandler<any>) {
    this._stringsEventDispatcher.add(event);
  }

  public static removeOnStringsUpdated(event: EventHandler<any>) {
    this._stringsEventDispatcher.remove(event);
  }

  public static onDirectionUpdated(event: EventHandler<any>) {
    this.initDirection();
    this._directionEventDispatcher.add(event);
  }

  public static removeOnDirectionUpdated(event: EventHandler<any>) {
    this.mutationObserver.disconnect();
    this._directionEventDispatcher.remove(event);
  }

  private static _isDirectionInit = false;

  /**
   * Checks for direction setup and adds mutationObserver
   *
   * @private
   * @static
   * @returns
   * @memberof LocalizationHelper
   */
  private static initDirection() {
    if (this._isDirectionInit) {
      return;
    }

    this._isDirectionInit = true;
    this.mutationObserver = new MutationObserver(function(mutations) {
      mutations.forEach(function(mutation) {
        if (mutation.attributeName == 'dir') {
          this._directionEventDispatcher.fire(null);
        }
      });
    });
    this.mutationObserver.observe(document.body, { childList: true });
  }

  /**
   * Provided helper method to determine localized or defaultString for specific string is returned
   *
   * @static updateStringsForTag
   * @param {string} tagName
   * @param  stringsObj
   * @returns
   * @memberof LocalizationHelper
   */
  public static updateStringsForTag(tagName: string, stringObj) {
    tagName = tagName.toLowerCase();

    if (tagName.startsWith('mgt-')) {
      tagName = tagName.substring(4);
    }

    if (this._strings && stringObj) {
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
