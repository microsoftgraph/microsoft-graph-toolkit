/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { styles } from './mgt-arrow-options-css';

/*
  Ok, the name here deserves a bit of explanation,
  This component originally had a built-in arrow icon,
  The problem came when you wanted to use a different symbol,
  So the arrow was removed, but the name was already set everywhere.
  - benotter
 */
@customElement('mgt-arrow-options')
export class MgtArrowOptions extends LitElement {
  public static get styles() {
    return styles;
  }

  @property({ type: Boolean }) public open: boolean = false;
  @property({ type: String }) public value: string = '';
  @property({ type: Object }) public options: { [name: string]: (e: MouseEvent) => any | void } = {};

  private _clickHandler: (e: MouseEvent) => void | any;

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

  public onHeaderClick(e: MouseEvent) {
    let keys = Object.keys(this.options);
    if (keys.length > 1) {
      e.preventDefault();
      e.stopPropagation();
      this.open = !this.open;
    }
  }

  public render() {
    return html`
      <span class="Header" @click=${e => this.onHeaderClick(e)}>
        <span class="CurrentValue">${this.value}</span>
      </span>
      <div class=${classMap({ Menu: true, Open: this.open, Closed: !this.open })}>
        ${this.getMenuOptions()}
      </div>
    `;
  }

  private getMenuOptions() {
    let keys = Object.keys(this.options);
    let funcs = this.options;

    return keys.map(
      opt => html`
        <div
          class="MenuOption"
          @click="${(e: MouseEvent) => {
            this.open = false;
            funcs[opt](e);
          }}"
        >
          <span class=${classMap({ MenuOptionCheck: true, CurrentValue: this.value === opt })}>
            \uE73E
          </span>
          <span class="MenuOptionName">${opt}</span>
        </div>
      `
    );
  }
}
