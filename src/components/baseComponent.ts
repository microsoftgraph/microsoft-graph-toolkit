/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, LitElement } from 'lit-element';
/**
 * BaseComponent extends LitElement including ShadowRoot toggle and fireCustomEvent features
 *
 * @export  MgtBaseComponent
 * @abstract
 * @class MgtBaseComponent
 * @extends {LitElement}
 */
export abstract class MgtBaseComponent extends LitElement {
  /**
   * Get ShadowRoot toggle, returns value of _useShadowRoot
   *
   * @static _useShadowRoot
   * @memberof MgtBaseComponent
   */
  public static get useShadowRoot() {
    return this._useShadowRoot;
  }

  /**
   * Set ShadowRoot toggle value
   *
   * @static _useShadowRoot
   * @memberof MgtBaseComponent
   */
  public static set useShadowRoot(value: boolean) {
    this._useShadowRoot = value;
  }

  private static _useShadowRoot: boolean = true;

  constructor() {
    super();
    if (this.isShadowRootDisabled()) {
      // tslint:disable-next-line: no-string-literal
      this['_needsShimAdoptedStyleSheets'] = true;
    }
  }

  /**
   * Recieve ShadowRoot Disabled value
   *
   * @returns boolean _useShadowRoot value
   * @memberof MgtBaseComponent
   */
  public isShadowRootDisabled() {
    return !MgtBaseComponent._useShadowRoot || !(this.constructor as typeof MgtBaseComponent)._useShadowRoot;
  }

  /**
   * helps facilitate creation of events across components
   *
   * @protected
   * @param {string} eventName name given to specific event
   * @param {*} [detail] optional any value to dispatch with event
   * @returns {boolean}
   * @memberof MgtBaseComponent
   */
  protected fireCustomEvent(eventName: string, detail?: any): boolean {
    const event = new CustomEvent(eventName, {
      bubbles: false,
      cancelable: true,
      detail
    });
    return this.dispatchEvent(event);
  }
  /**
   * method to create ShadowRoot if disabled flag isn't present
   *
   * @protected
   * @returns boolean
   * @memberof MgtBaseComponent
   */
  protected createRenderRoot() {
    return this.isShadowRootDisabled() ? this : super.createRenderRoot();
  }
}
