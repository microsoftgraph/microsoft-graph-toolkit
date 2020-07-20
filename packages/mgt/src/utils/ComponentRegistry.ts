/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement } from 'lit-element';

/**
 * Manages the registration of custom elements.
 *
 * @export
 * @class ComponentRegistry
 */
export class ComponentRegistry {
  /**
   * Maintains the collection of custom elements available in the toolkit.
   *
   * @readonly
   * @static
   * @type {object}
   * @memberof ComponentRegistry
   */
  public static get scopedElements(): object {
    return this._scopedElements;
  }

  /**
   * Register a custom component for use in the DOM.
   *
   * @static
   * @param {string} tagName
   * @param {*} component
   * @memberof ComponentRegistry
   */
  public static register(tagName: string, component: any): void {
    // Add the component to the scopedElements object.
    ComponentRegistry._scopedElements[tagName] = component;

    // Check for existing registration
    if (!customElements.get(tagName)) {
      // Register component
      customElement(tagName)(component);
    }
  }

  private static _scopedElements: object = {};
}
