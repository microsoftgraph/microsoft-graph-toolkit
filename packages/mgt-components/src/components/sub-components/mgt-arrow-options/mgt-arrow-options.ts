/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { MgtBaseComponent } from '@microsoft/mgt-element';
import { styles } from './mgt-arrow-options-css';

/*
  Ok, the name here deserves a bit of explanation,
  This component originally had a built-in arrow icon,
  The problem came when you wanted to use a different symbol,
  So the arrow was removed, but the name was already set everywhere.
  - benotter
 */

/**
 * Custom Component used to handle an arrow rendering for TaskGroups utilized in the task component.
 *
 * @export MgtArrowOptions
 * @class MgtArrowOptions
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-arrow-options')
export class MgtArrowOptions extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  public static get styles() {
    return styles;
  }

  /**
   * Determines if header menu is rendered or hidden.
   *
   * @type {boolean}
   * @memberof MgtArrowOptions
   */
  @property({ type: Boolean }) public open: boolean;

  /**
   * Title of chosen TaskGroup.
   *
   * @type {string}
   * @memberof MgtArrowOptions
   */
  @property({ type: String }) public value: string;

  /**
   * Menu options to be rendered with an attached MouseEvent handler for expansion of details
   *
   * @type {object}
   * @memberof MgtArrowOptions
   */
  @property({ type: Object }) public options: { [name: string]: (e: MouseEvent) => any | void };

  private _clickHandler: (e: MouseEvent) => void | any;

  constructor() {
    super();
    this.value = '';
    this.options = {};
    this._clickHandler = (e: MouseEvent) => (this.open = false);
  }

  // tslint:disable-next-line: completed-docs
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this._clickHandler);
  }

  // tslint:disable-next-line: completed-docs
  public disconnectedCallback() {
    window.removeEventListener('click', this._clickHandler);
    super.disconnectedCallback();
  }

  /**
   * Handles clicking for header menu, utilizing boolean switch open
   *
   * @param {MouseEvent} e attaches to Header to open menu
   * @memberof MgtArrowOptions
   */
  public onHeaderClick(e: MouseEvent) {
    const keys = Object.keys(this.options);
    if (keys.length > 1) {
      e.preventDefault();
      e.stopPropagation();
      this.open = !this.open;
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
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
    const keys = Object.keys(this.options);
    const funcs = this.options;

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
