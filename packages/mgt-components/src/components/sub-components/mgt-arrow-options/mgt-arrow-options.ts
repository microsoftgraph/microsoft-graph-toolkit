/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { property } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { MgtBaseComponent, customElement } from '@microsoft/mgt-element';
import { styles } from './mgt-arrow-options-css';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { fluentMenu, fluentMenuItem } from '@fluentui/web-components';
registerFluentComponents(fluentMenu, fluentMenuItem);

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
@customElement('arrow-options')
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
   * Menu options to be rendered with an attached UIEvent handler for expansion of details
   *
   * @type {object}
   * @memberof MgtArrowOptions
   */
  @property({ type: Object }) public options: { [name: string]: (e: UIEvent) => any | void };

  private _clickHandler: (e: UIEvent) => void | any;

  constructor() {
    super();
    this.value = '';
    this.options = {};
    this._clickHandler = () => (this.open = false);
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
   * Handles clicking for header menu, utilizing boolean switch open
   *
   * @param {UIEvent} e attaches to Header to open menu
   * @memberof MgtArrowOptions
   */
  public onHeaderClick = (e: UIEvent) => {
    const keys = Object.keys(this.options);
    if (keys.length > 1) {
      e.preventDefault();
      e.stopPropagation();
      this.open = !this.open;
    }
  };

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    return html`
      <a class="Header" @click=${this.onHeaderClick} href="javascript:void">${this.value}</a>
      <fluent-menu
        class=${classMap({ menu: true, open: this.open, closed: !this.open })}>
          ${this.getMenuOptions()}
      </fluent-menu>`;
  }

  private getMenuOptions() {
    const keys = Object.keys(this.options);

    return keys.map((opt: string) => {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-return
      const clickFn = (e: MouseEvent) => {
        this.open = false;
        this.options[opt](e);
      };

      const keyDownFn = (e: KeyboardEvent) => {
        if (e.key === 'Enter') {
          this.open = false;
          this.options[opt](e);
        }
      };
      return html`
          <fluent-menu-item
            @click=${clickFn}
            @keydown=${keyDownFn}>
              ${opt}
          </fluent-menu-item>`;
    });
  }
}
