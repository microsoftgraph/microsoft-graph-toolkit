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

  private static _disableAllShadowRoots: boolean = false;
  public static get disableAllShadowRoots() {
    return this._disableAllShadowRoots;
  }
  public static set disableAllShadowRoots(value: boolean) {
    this._disableAllShadowRoots = value;
  }

  constructor() {
    super();
    if ((this.constructor as typeof MgtBaseComponent)._disableAllShadowRoots)
      this['_needsShimAdoptedStyleSheets'] = true;
  }

  protected createRenderRoot() {
    return MgtBaseComponent._disableAllShadowRoots ||
      (this.constructor as (typeof MgtBaseComponent))._disableAllShadowRoots
      ? this
      : super.createRenderRoot();
  }
}

if (window && !window[MgtBaseComponent.name]) window[MgtBaseComponent.name] = MgtBaseComponent;
