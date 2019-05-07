/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { styles } from './mgt-dot-options-css';

@customElement('mgt-dot-options')
export class MgtDotOptions extends LitElement {
  public static get styles() {
    return styles;
  }

  @property({ type: Boolean }) public open: boolean = false;
  @property({ type: Object }) public options: { [option: string]: (e: MouseEvent) => void | any } = null;

  private _clickHandler: (e: MouseEvent) => void | any = null;

  public constructor() {
    super();
    this._clickHandler = (e: MouseEvent) => (this.open = false);
  }

  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this._clickHandler);
  }

  public disconnectedCallback() {
    window.removeEventListener('click', this._clickHandler);
    super.disconnectedCallback();
  }

  private onDotClick(e: MouseEvent) {
    e.preventDefault();
    e.stopPropagation();

    this.open = !this.open;
  }

  public render() {
    return html`
      <div class=${classMap({ DotMenu: true, Open: this.open })} @click=${e => this.onDotClick(e)}>
        <span class="DotIcon">\uE712</span>
        <div class="Menu">
          ${Object.keys(this.options).map(prop => this.getMenuOption(prop, this.options[prop]))}
        </div>
      </div>
    `;
  }

  public getMenuOption(name: string, click: (e: MouseEvent) => void | any) {
    return html`
      <div
        class="DotItem"
        @click="${e => {
          e.preventDefault();
          e.stopPropagation();
          click(e);
          this.open = false;
        }}"
      >
        <span class="DotItemName">
          ${name}
        </span>
      </div>
    `;
  }
}
