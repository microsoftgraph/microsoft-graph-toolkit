/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, LitElement, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { styles } from './mgt-flyout-css';

/**
 * A component to create flyout anchored to an element
 *
 * @export
 * @class MgtFlyout
 * @extends {LitElement}
 */
@customElement('mgt-flyout')
export class MgtFlyout extends LitElement {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Gets or sets whether the flyout is visible
   *
   * @type {string}
   * @memberof MgtComponent
   */
  @property({
    attribute: 'isOpen',
    type: Boolean
  })
  public isOpen: boolean;

  private renderedOnce = false;

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  public firstUpdated() {
    this.addEventListener('updated', e => {
      this.updateFlyout();
    });
  }

  /**
   * Invoked whenever the element is updated. Implement to perform
   * post-updating tasks via DOM APIs, for example, focusing an element.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param changedProperties Map of changed properties with old values
   */
  protected updated(changedProps: PropertyValues) {
    super.updated(changedProps);
    this.updateFlyout();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html`
      <div class="root">
        <div class="anchor">
          <slot></slot>
        </div>
        ${this.renderFlyout()}
      </div>
    `;
  }

  private renderFlyout() {
    if (!this.isOpen && !this.renderedOnce) {
      return;
    }

    this.renderedOnce = true;

    const classes = {
      flyout: true,
      visible: this.isOpen
    };

    return html`
      <div class=${classMap(classes)}>
        <slot name="flyout"></slot>
      </div>
    `;
  }

  private updateFlyout() {
    const anchor = this.renderRoot.querySelector('.anchor');
    const flyout = this.renderRoot.querySelector('.flyout') as HTMLElement;
    if (flyout && anchor) {
      let left: number;
      let bottom: number;
      let top: number;

      if (this.isOpen) {
        const flyoutRect = flyout.getBoundingClientRect();
        const anchorRect = anchor.getBoundingClientRect();

        // normalize flyoutrect since we could have moved it before
        // need to know where would it render, not where it renders
        const flyoutTop = anchorRect.bottom;
        const flyoutLeft = anchorRect.left;
        const flyoutRight = flyoutLeft + flyoutRect.width;
        const flyoutBottom = flyoutTop + flyoutRect.height;

        const windowWidth =
          window.innerWidth && document.documentElement.clientWidth
            ? Math.min(window.innerWidth, document.documentElement.clientWidth)
            : window.innerWidth || document.documentElement.clientWidth;

        const windowHeight =
          window.innerHeight && document.documentElement.clientHeight
            ? Math.min(window.innerHeight, document.documentElement.clientHeight)
            : window.innerHeight || document.documentElement.clientHeight;

        if (flyoutRect.width > windowWidth) {
          // page width is smaller than flyout, render all the way to the left
          left = -flyoutLeft;
        } else if (anchorRect.width >= flyoutRect.width) {
          // anchor is large than flyout, render aligned to anchor
          left = 0;
        } else {
          const centerOffset = flyoutRect.width / 2 - anchorRect.width / 2;

          if (flyoutLeft - centerOffset < 0) {
            // centered flyout is off screen to the left, render on the left edge
            left = -flyoutLeft;
          } else if (flyoutRight - centerOffset > windowWidth) {
            // centered flyout is off screen to the right, render on the right edge
            left = -(flyoutRight - windowWidth);
          } else {
            // render centered
            left = -centerOffset;
          }
        }

        if (flyoutRect.height > windowHeight || (windowHeight < flyoutBottom && anchorRect.top < flyoutRect.height)) {
          top = -flyoutTop + anchorRect.height;
        } else if (windowHeight < flyoutBottom) {
          bottom = anchorRect.height;
        }
      }

      flyout.style.left = typeof left !== 'undefined' ? `${left}px` : '';
      flyout.style.bottom = typeof bottom !== 'undefined' ? `${bottom}px` : '';
      flyout.style.top = typeof top !== 'undefined' ? `${top}px` : '';
    }
  }
}
