/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { MgtBaseComponent } from '@microsoft/mgt-element';
import { styles } from './mgt-dot-options-css';
/**
 * Custom Component used to handle an arrow rendering for TaskGroups utilized in the task component.
 *
 * @export MgtDotOptions
 * @class MgtDotOptions
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-dot-options')
export class MgtDotOptions extends MgtBaseComponent {
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
   * @memberof MgtDotOptions
   */
  @property({ type: Boolean }) public open: boolean;

  /**
   * Menu options to be rendered with an attached MouseEvent handler for expansion of details
   *
   * @memberof MgtDotOptions
   */
  @property({ type: Object }) public options: { [option: string]: (e: MouseEvent) => void | any };

  private _clickHandler: (e: MouseEvent) => void | any = null;

  constructor() {
    super();
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
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
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
  /**
   * Used by the render method to attach click handler to each dot item
   *
   * @param {string} name
   * @param {((e: MouseEvent) => void | any)} click
   * @returns
   * @memberof MgtDotOptions
   */
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

  private onDotClick(e: MouseEvent) {
    e.preventDefault();
    e.stopPropagation();

    this.open = !this.open;
  }
}
