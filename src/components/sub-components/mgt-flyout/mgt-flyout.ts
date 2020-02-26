/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, LitElement, property, PropertyValues } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { getSegmentAwareWindow, isWindowSegmentAware, IWindowSegment } from '../../../utils/DualScreenHelpers';
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
  private openLeft: boolean = false;

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
        // Pre-calculations
        const documentClientHeight = document.documentElement.clientHeight;
        const innerHeight =
          window.innerHeight && documentClientHeight
            ? Math.min(window.innerHeight, documentClientHeight)
            : window.innerHeight || documentClientHeight;

        const documentClientWidth = document.documentElement.clientWidth;
        const innerWidth =
          window.innerWidth && documentClientWidth
            ? Math.min(window.innerWidth, documentClientWidth)
            : window.innerWidth || documentClientWidth;

        const anchorRect = anchor.getBoundingClientRect();
        const anchorCenterX = anchorRect.left + anchorRect.width / 2;
        const anchorCenterY = anchorRect.top + anchorRect.height / 2;

        const windowRect: IWindowSegment = {
          height: 0,
          left: 0,
          top: 0,
          width: 0
        };

        if (isWindowSegmentAware()) {
          const segmentAwareWindow = getSegmentAwareWindow();
          const screenSegments = segmentAwareWindow.getWindowSegments();

          let anchorSegment: IWindowSegment;
          for (const segment of screenSegments) {
            if (anchorCenterX >= segment.left && anchorCenterY >= segment.top) {
              anchorSegment = segment;
            }
          }

          windowRect.left = anchorSegment.left;
          windowRect.top = anchorSegment.top;
          windowRect.width = anchorSegment.width;
          windowRect.height = anchorSegment.height;
        } else {
          windowRect.height = innerHeight;
          windowRect.width = innerWidth;
        }

        // normalize flyoutrect since we could have moved it before
        // need to know where would it render, not where it renders
        const flyoutWidth = flyout.scrollWidth;
        const flyoutHeight = flyout.scrollHeight;
        const flyoutTop = anchorRect.bottom;
        const flyoutLeft = anchorRect.left;
        const flyoutRight = flyoutLeft + flyoutWidth;
        const flyoutBottom = flyoutTop + flyoutHeight;

        // Rules of alignment:
        // 1. Center if possible
        // 2. Don't cross screen boundaries.

        if (flyoutWidth > windowRect.width) {
          // page width is smaller than flyout, render all the way to the left
          left = -anchorRect.left;
        } else if (Math.floor(anchorRect.width) >= flyoutWidth) {
          // anchor is larger than flyout, render aligned to anchor
          left = anchorRect.left;
        } else {
          const centerOffset = flyoutWidth / 2 - anchorRect.width / 2;

          if (anchorRect.left - centerOffset < windowRect.left) {
            // centered flyout is off screen to the left, render on the left edge
            left = windowRect.left - anchorRect.left;
          } else if (flyoutRight - centerOffset > windowRect.width) {
            // centered flyout is off screen to the right, render on the right edge
            left =
              -centerOffset * 2 -
              (anchorRect.left +
                Math.floor(anchorRect.width) -
                Math.min(windowRect.left + windowRect.width, innerWidth)) -
              1;
          } else {
            // render centered
            left = -centerOffset;
          }
        }

        if (flyoutHeight > windowRect.height || (windowRect.height < flyoutBottom && anchorRect.top < flyoutHeight)) {
          top = -flyoutTop + anchorRect.height;
        } else if (windowRect.height < flyoutBottom) {
          bottom = anchorRect.height;
        }
      }

      flyout.style.left = typeof left !== 'undefined' ? `${left}px` : '';
      flyout.style.bottom = typeof bottom !== 'undefined' ? `${bottom}px` : '';
      flyout.style.top = typeof top !== 'undefined' ? `${top}px` : '';
    }
  }
}
