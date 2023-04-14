/* eslint-disable @typescript-eslint/tslint/config */
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { EventDispatcher, EventHandler } from './EventDispatcher';

/**
 * Valid values for the direction attribute
 */
export type Direction = 'ltr' | 'rtl' | 'auto';

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

  private static mutationObserver: MutationObserver;

  public static get strings() {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-return
    return this._strings;
  }

  /**
   * Set strings to be localized
   *
   * @static
   * @memberof LocalizationHelper
   */
  public static set strings(value: any) {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
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
  public static getDocumentDirection(): 'rtl' | 'ltr' | 'auto' {
    // Re-set the dir to ltr if the dir attribute is already loaded and the first two options
    // are returning null values.
    const parsed = document.body?.getAttribute('dir') || document.documentElement?.getAttribute('dir');
    switch (parsed) {
      case 'rtl':
        return 'rtl';
      case 'auto':
        return 'auto';
      case 'ltr':
      default:
        return 'ltr';
    }
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
    this._directionEventDispatcher.add(event);
    this.initDirection();
  }

  public static removeOnDirectionUpdated(event: EventHandler<any>) {
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
    this.mutationObserver = new MutationObserver(mutations => {
      mutations.forEach(mutation => {
        if (mutation.attributeName === 'dir') {
          this._directionEventDispatcher.fire(null);
        }
      });
    });
    const options = { attributes: true, attributeFilter: ['dir'] };
    this.mutationObserver.observe(document.body, options);
    this.mutationObserver.observe(document.documentElement, options);
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
  public static updateStringsForTag(tagName: string, stringObj: Record<string, any>) {
    tagName = tagName.toLowerCase();

    if (tagName.startsWith('mgt-')) {
      tagName = tagName.substring(4);
    }

    if (this._strings && stringObj) {
      // check for top level strings, applied per component, overridden by specific component def
      for (const prop of Object.entries(stringObj)) {
        // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
        if (this._strings[prop[0]]) {
          // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
          stringObj[prop[0]] = this._strings[prop[0]];
        }
      }
      // strings defined component specific
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
      if (this._strings._components && this._strings._components[tagName]) {
        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
        const strings: any = this._strings._components[tagName];
        // eslint-disable-next-line @typescript-eslint/no-unsafe-argument
        for (const key of Object.keys(strings)) {
          if (stringObj[key]) {
            // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-member-access
            stringObj[key] = strings[key];
          }
        }
      }
    }

    return stringObj;
  }
}
