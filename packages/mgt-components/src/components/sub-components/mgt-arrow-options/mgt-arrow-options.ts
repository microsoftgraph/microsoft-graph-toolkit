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
import { fluentMenu, fluentMenuItem, fluentButton } from '@fluentui/web-components';
registerFluentComponents(fluentMenu, fluentMenuItem, fluentButton);

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
 * @cssprop --arrow-options-left {Length} The distance of the dropdown menu from the left in absolute position. Default is 0.
 * @cssprop --arrow-options-button-background-color {Color} The background color of the arrow options button.
 * @cssprop --arrow-options-button-font-size {Length} The font size of the button text. Default is large.
 * @cssprop --arrow-options-button-font-weight {Length} The font weight of the button text. Default is 600.
 * @cssprop --arrow-options-button-font-color {Color} The font color of the text in the button.
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
    window.addEventListener('onblur', () => (this.open = false));
  }

  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this._clickHandler);
  }

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
  public onHeaderClick = (e: MouseEvent) => {
    const keys = Object.keys(this.options);
    if (keys.length > 1) {
      e.preventDefault();
      e.stopPropagation();
      this.open = !this.open;
    }
  };

  /**
   * Handles key down presses done on the header element.
   *
   * @param {KeyboardEvent} e
   */
  private onHeaderKeyDown = (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      this.open = !this.open;

      // Manually adding the 'open' class to display the menu because
      // by the time I set the first element's focus, the classes are not
      // updated and that has no effect. You can't set focus on elements
      // that have no display.
      const fluentMenuEl: HTMLElement = this.renderRoot.querySelector('fluent-menu');
      if (fluentMenuEl) {
        fluentMenuEl.classList.remove('closed');
        fluentMenuEl.classList.add('open');
      }

      const header: HTMLButtonElement = e.target as HTMLButtonElement;
      if (header) {
        const firstMenuItem: HTMLElement = this.renderRoot.querySelector("fluent-menu-item[tabindex='0']");
        if (firstMenuItem) {
          header.blur();
          firstMenuItem.focus();
        }
      }
    }
  };

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    return html`
      <fluent-button
        class="header"
        @click=${this.onHeaderClick}
        @keydown=${this.onHeaderKeyDown}
        appearance="lightweight">
          ${this.value}
      </fluent-button>
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
        const header: HTMLButtonElement = this.renderRoot.querySelector<HTMLButtonElement>('.header');
        if (e.key === 'Enter') {
          this.open = false;
          this.options[opt](e);
          header.focus();
        } else if (e.key === 'Tab') {
          this.open = false;
        } else if (e.key === 'Escape') {
          this.open = false;
          if (header) {
            header.focus();
          }
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
