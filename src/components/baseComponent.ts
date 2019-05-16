/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, html } from 'lit-element';

export abstract class MgtBaseComponent extends LitElement {
  protected fireCustomEvent(eventName: string, detail?: any): boolean {
    let event = new CustomEvent(eventName, {
      cancelable: true,
      bubbles: false,
      detail: detail
    });
    return this.dispatchEvent(event);
  }

  private static _useShadowRoot: boolean = true;
  public static get useShadowRoot() {
    return this._useShadowRoot;
  }
  public static set useShadowRoot(value: boolean) {
    this._useShadowRoot = value;
  }

  constructor() {
    super();
    if (this.isShadowRootDisabled()) this['_needsShimAdoptedStyleSheets'] = true;
  }

  protected createRenderRoot() {
    return this.isShadowRootDisabled() ? this : super.createRenderRoot();
  }

  public isShadowRootDisabled() {
    return !MgtBaseComponent._useShadowRoot || !(this.constructor as typeof MgtBaseComponent)._useShadowRoot;
  }
}
