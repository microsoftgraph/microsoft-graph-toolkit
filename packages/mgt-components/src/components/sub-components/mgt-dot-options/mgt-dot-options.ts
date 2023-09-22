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
import { strings } from './strings';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { styles } from './mgt-dot-options-css';

registerFluentComponents(fluentMenu, fluentMenuItem, fluentButton);

/**
 * Defines the event functions passed to the option item.
 */
type MenuOptionEventFunction = (e: Event) => void;

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
   * Strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtDotOptions
   */
  protected get strings() {
    return strings;
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
  @property({ type: Object }) public options: Record<string, (e: Event) => void>;

  private readonly _clickHandler = () => (this.open = false);

  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this._clickHandler);
  }

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
      <fluent-button
        appearance="stealth"
        aria-label=${this.strings.dotOptionsTitle}
        @click=${this.onDotClick}
        @keydown=${this.onDotKeydown}
        class="dot-icon">\uE712</fluent-button>
      <fluent-menu class=${classMap({ menu: true, open: this.open })}>
        ${menuOptions.map(opt => this.getMenuOption(opt, this.options[opt]))}
      </fluent-menu>`;
  }

  private readonly handleItemClick = (e: MouseEvent, fn: MenuOptionEventFunction) => {
    e.preventDefault();
    e.stopPropagation();
    fn(e);
    this.open = false;
  };

  private readonly handleItemKeydown = (e: KeyboardEvent, fn: MenuOptionEventFunction) => {
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
  public getMenuOption(name: string, clickFn: (e: Event) => void) {
    return html`
      <fluent-menu-item
        @click=${(e: MouseEvent) => this.handleItemClick(e, clickFn)}
        @keydown=${(e: KeyboardEvent) => this.handleItemKeydown(e, clickFn)}>
          ${name}
      </fluent-menu-item>`;
  }

  private readonly onDotClick = (e: MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    this.open = !this.open;
  };

  private readonly onDotKeydown = (e: KeyboardEvent) => {
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
