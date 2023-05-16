/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fluentMenu, fluentMenuItem, fluentButton } from '@fluentui/web-components';
import { MgtBaseComponent, customElement } from '@microsoft/mgt-element';
import { html } from 'lit';
import { property } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { styles } from './mgt-dot-options-css';

registerFluentComponents(fluentMenu, fluentMenuItem, fluentButton);

/**
 * Defines the event functions passed to the option item.
 */
type MenuOptionEventFunction = (e: Event) => void | any;

/**
 * Custom Component used to handle an arrow rendering for TaskGroups utilized in the task component.
 *
 * @export MgtDotOptions
 * @class MgtDotOptions
 * @extends {MgtBaseComponent}
 */
@customElement('dot-options')
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
  @property({ type: Object }) public options: { [option: string]: (e: Event) => void | any };

  private _clickHandler: (e: MouseEvent) => void | any = null;

  constructor() {
    super();
    this._clickHandler = (e: MouseEvent) => (this.open = false);
  }

  // eslint-disable-next-line @typescript-eslint/tslint/config
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this._clickHandler);
  }

  // eslint-disable-next-line @typescript-eslint/tslint/config
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
    const menuOptions = Object.keys(this.options);

    return html`
      <div tabindex="0" class=${classMap({ 'dot-menu': true, open: this.open })}
        @click=${this.onDotClick}
        @keydown=${this.onDotKeydown}>
        <span class="dot-icon">\uE712</span>
        <div tabindex="0" class="menu">
          ${Object.keys(this.options).map(prop => this.getMenuOption(prop, this.options[prop]))}
        </div>
      </div>
    `;
  }

  private handleItemClick = (e: MouseEvent, fn: MenuOptionEventFunction) => {
    e.preventDefault();
    e.stopPropagation();
    fn(e);
    this.open = false;
  };

  private handleItemKeydown = (e: KeyboardEvent, fn: MenuOptionEventFunction) => {
    this.handleKeydownMenuOption(e);
    fn(e);
    this.open = false;
  };

  /**
   * Used by the render method to attach click handler to each dot item
   *
   * @param {string} name
   * @param {MenuOptionEventFunction} clickFn
   * @returns
   * @memberof MgtDotOptions
   */
  public getMenuOption(name: string, clickFn: (e: Event) => void | any) {
    return html`
      <div
        class="dot-item"
        @click="${(e: Event) => {
          e.preventDefault();
          e.stopPropagation();
          clickFn(e);
          this.open = false;
        }}"
        @keydown="${(e: KeyboardEvent) => {
          if (e.code === 'Enter' || e.code === 'Space') {
            e.preventDefault();
            e.stopPropagation();
            clickFn(e);
            this.open = false;
          }
        }}"
      >
        <span class="dot-item-name">
          ${name}
      </fluent-menu-item>`;
  }

  private onDotClick = (e: MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    this.open = !this.open;
  };

  private onDotKeydown = (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();

      this.open = !this.open;
    }
  };

  private handleKeydownMenuOption(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
    }
  }
}
