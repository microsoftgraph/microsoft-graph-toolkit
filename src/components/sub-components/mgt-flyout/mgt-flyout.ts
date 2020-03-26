/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, LitElement, property, PropertyValues, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { getSegmentAwareWindow, isWindowSegmentAware, IWindowSegment } from '../../../utils/WindowSegmentHelpers';
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
   * Gets or sets whether the flyout is light dismissable.
   *
   * @type {boolean}
   * @memberof MgtFlyout
   */
  @property({
    attribute: 'light-dismiss',
    type: Boolean
  })
  public isLightDismiss: boolean;

  /**
   * Gets or sets whether the flyout is visible
   *
   * @type {string}
   * @memberof MgtFlyout
   */
  @property({
    attribute: 'isOpen',
    type: Boolean
  })
  public get isOpen(): boolean {
    return this._isOpen;
  }
  public set isOpen(value: boolean) {
    if (this._isOpen === value) {
      return;
    }
    this._isOpen = value;
    this.setupWindowEvents(value);
    this.requestUpdate('isOpen');
    this.dispatchEvent(new Event(value ? 'opened' : 'closed'));
  }

  private _isOpen: boolean;

  constructor() {
    super();
    this.handleWindowEvent = this.handleWindowEvent.bind(this);
    this.handleWindowClickEvent = this.handleWindowClickEvent.bind(this);
  }

  /**
   * Show the flyout.
   */
  public open(): void {
    this.isOpen = true;
  }

  /**
   * Close the flyout.
   */
  public close(): void {
    this.isOpen = false;
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtFlyout
   */
  public disconnectedCallback() {
    this.setupWindowEvents(false);
    super.disconnectedCallback();
  }

  /**
   * Invoked whenever the element is updated. Implement to perform
   * post-updating tasks via DOM APIs, for example, focusing an element.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * @param changedProperties Map of changed properties with old values
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
    const anchorTemplate = this.renderAnchor();
    const flyoutTemplate = this.isOpen ? this.renderFlyout() : null;

    const flyoutClasses = {
      flyout: true,
      visible: this.isOpen
    };

    return html`
      <div class="root">
        <div class="anchor">
          ${anchorTemplate}
        </div>
        <div class=${classMap(flyoutClasses)}>
          ${flyoutTemplate}
        </div>
      </div>
    `;
  }

  /**
   * Renders the anchor content.
   *
   * @protected
   * @returns
   * @memberof MgtFlyout
   */
  protected renderAnchor(): TemplateResult {
    return html`
      <slot></slot>
    `;
  }

  /**
   * Renders the flyout.
   */
  protected renderFlyout(): TemplateResult {
    return html`
      <slot name="flyout"></slot>
    `;
  }

  private updateFlyout() {
    if (!this.isOpen) {
      return;
    }

    const anchor = this.renderRoot.querySelector('.anchor');
    const flyout = this.renderRoot.querySelector('.flyout') as HTMLElement;
    if (flyout && anchor) {
      let left: number;
      let bottom: number;
      let top: number;

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
        const centerOffset = Math.floor(flyoutWidth / 2 - anchorRect.width / 2);

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

      flyout.style.left = typeof left !== 'undefined' ? `${left}px` : '';
      flyout.style.bottom = typeof bottom !== 'undefined' ? `${bottom}px` : '';
      flyout.style.top = typeof top !== 'undefined' ? `${top}px` : '';
    }
  }

  private setupWindowEvents(isOpen: boolean): void {
    if (isOpen && this.isLightDismiss) {
      window.addEventListener('click', this.handleWindowClickEvent, true);
      window.addEventListener('resize', this.handleWindowEvent);
    } else {
      window.removeEventListener('click', this.handleWindowClickEvent, true);
      window.removeEventListener('resize', this.handleWindowEvent);
    }
  }

  private handleWindowEvent(e: Event): void {
    this.close();
  }

  private handleWindowClickEvent(e: Event): void {
    const flyout = this.renderRoot.querySelector('.flyout');

    if (flyout) {
      // IE
      if (!e.composedPath) {
        let currentElem = e.target as HTMLElement;
        while (currentElem) {
          currentElem = currentElem.parentElement;
          if (currentElem === flyout || (e.type === 'pointerdown' && currentElem === this)) {
            return;
          }
        }
      } else {
        const path = e.composedPath();
        if (path.includes(flyout) || (e.type === 'pointerdown' && path.includes(this))) {
          return;
        }
      }
    }

    this.handleWindowEvent(e);
  }
}
