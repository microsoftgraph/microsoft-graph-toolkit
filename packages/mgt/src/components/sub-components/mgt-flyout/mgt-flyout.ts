/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, PropertyValues, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { getSegmentAwareWindow, isWindowSegmentAware, IWindowSegment } from '../../../utils/WindowSegmentHelpers';
import { styles } from './mgt-flyout-css';
import { MgtBaseComponent } from '@microsoft/mgt-element/';

/**
 * A component to create flyout anchored to an element
 *
 * @export
 * @class MgtFlyout
 * @extends {LitElement}
 */
@customElement('mgt-flyout')
export class MgtFlyout extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Gets or sets whether the flyout is light dismissible.
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
    const oldValue = this._isOpen;
    if (oldValue === value) {
      return;
    }

    this._isOpen = value;

    window.requestAnimationFrame(() => {
      this.setupWindowEvents(this.isOpen);
      const flyout = this._flyout;
      if (!this.isOpen && flyout) {
        // reset style for next update
        flyout.style.width = null;
        flyout.style.setProperty('--mgt-flyout-set-width', null);
        flyout.style.height = null;
        flyout.style.top = null;
        flyout.style.left = null;
        flyout.style.bottom = null;
      }
    });

    this.requestUpdate('isOpen', oldValue);
    this.dispatchEvent(new Event(value ? 'opened' : 'closed'));
  }

  // Minimum distance to render from window edge
  private _edgePadding: number = 24;

  // if the flyout is opened once, this will keep the flyout in the dom
  private _renderedOnce = false;

  private _isRTL = false;

  private get _flyout(): HTMLElement {
    return this.renderRoot.querySelector('.flyout');
  }
  private get _anchor(): HTMLElement {
    return this.renderRoot.querySelector('.anchor');
  }

  private _isOpen: boolean;

  constructor() {
    super();

    this.handleWindowEvent = this.handleWindowEvent.bind(this);
    this.handleResize = this.handleResize.bind(this);
    this.handleKeyUp = this.handleKeyUp.bind(this);
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

    window.requestAnimationFrame(() => {
      this.updateFlyout();
    });
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const flyoutClasses = {
      root: true,
      visible: this.isOpen,
      dir: this.direction
    };

    const anchorTemplate = this.renderAnchor();
    let flyoutTemplate = null;

    if (this.isOpen || this._renderedOnce) {
      this._renderedOnce = true;
      flyoutTemplate = html`
        <div class="flyout" @wheel=${this.handleFlyoutWheel}>
          ${this.renderFlyout()}
        </div>
      `;
    }

    return html`
      <div class=${classMap(flyoutClasses)}>
        <div class="anchor">
          ${anchorTemplate}
        </div>
        ${flyoutTemplate}
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

    const anchor = this._anchor;
    const flyout = this._flyout;

    if (flyout && anchor) {
      const windowWidth =
        window.innerWidth && document.documentElement.clientWidth
          ? Math.min(window.innerWidth, document.documentElement.clientWidth)
          : window.innerWidth || document.documentElement.clientWidth;

      const windowHeight =
        window.innerHeight && document.documentElement.clientHeight
          ? Math.min(window.innerHeight, document.documentElement.clientHeight)
          : window.innerHeight || document.documentElement.clientHeight;

      let left: number = 0;
      let bottom: number;
      let top: number = 0;
      let height: number;
      let width: number;

      const flyoutRect = flyout.getBoundingClientRect();
      const anchorRect = anchor.getBoundingClientRect();

      const windowRect: IWindowSegment = {
        height: windowHeight,
        left: 0,
        top: 0,
        width: windowWidth
      };

      if (isWindowSegmentAware()) {
        const segmentAwareWindow = getSegmentAwareWindow();
        const screenSegments = segmentAwareWindow.getWindowSegments();

        let anchorSegment: IWindowSegment;

        const anchorCenterX = anchorRect.left + anchorRect.width / 2;
        const anchorCenterY = anchorRect.top + anchorRect.height / 2;

        for (const segment of screenSegments) {
          if (anchorCenterX >= segment.left && anchorCenterY >= segment.top) {
            anchorSegment = segment;
            break;
          }
        }

        if (anchorSegment) {
          windowRect.left = anchorSegment.left;
          windowRect.top = anchorSegment.top;
          windowRect.width = anchorSegment.width;
          windowRect.height = anchorSegment.height;
        }
      }

      if (flyoutRect.width + 2 * this._edgePadding > windowRect.width) {
        if (flyoutRect.width > windowRect.width) {
          // flyout is wider than the window
          width = windowRect.width;
          left = 0;
        } else {
          // center in between
          left = (windowRect.width - flyoutRect.width) / 2;
        }
      } else if (anchorRect.left + flyoutRect.width + this._edgePadding > windowRect.width) {
        // it will render off screen to the right, move to the left
        left = anchorRect.left - (anchorRect.left + flyoutRect.width + this._edgePadding - windowRect.width);
      } else if (anchorRect.left < this._edgePadding) {
        // it will render off screen to the left, move to the right
        left = this._edgePadding;
      } else {
        left = anchorRect.left;
      }

      if (flyoutRect.height + 2 * this._edgePadding > windowRect.height) {
        if (flyoutRect.height >= windowRect.height) {
          height = windowRect.height;
          top = 0;
        } else {
          top = (windowRect.height - flyoutRect.height) / 2;
        }
      } else if (
        anchorRect.top + anchorRect.height + flyoutRect.height + this._edgePadding > windowRect.height &&
        anchorRect.top - flyoutRect.height - this._edgePadding > 0
      ) {
        if (windowRect.height - anchorRect.top + flyoutRect.height < 0) {
          bottom = windowRect.height - flyoutRect.height - this._edgePadding;
        } else {
          bottom = Math.max(windowRect.height - anchorRect.top, this._edgePadding);
        }
      } else {
        if (anchorRect.top + anchorRect.height + flyoutRect.height + this._edgePadding > windowRect.height) {
          // it will render offscreen bellow, move it up a bit
          top = windowRect.height - flyoutRect.height - this._edgePadding;
        } else {
          top = Math.max(anchorRect.top + anchorRect.height, this._edgePadding);
        }
      }

      if (this.direction === 'rtl') {
        if (left > 100) {
          //potentially anchored to right side (for non people-picker flyout)
          flyout.style.left = `${windowRect.width - left + flyoutRect.left - flyoutRect.width - 30}px`;
        }
      } else {
        flyout.style.left = `${left + windowRect.left}px`;
      }

      if (typeof bottom !== 'undefined') {
        flyout.style.top = 'unset';
        flyout.style.bottom = `${bottom}px`;
      } else {
        flyout.style.bottom = 'unset';
        flyout.style.top = `${top + windowRect.top}px`;
      }

      flyout.style.height = height ? `${height}px` : null;

      if (width) {
        // if we had to set the width, recalculate since the height could have changed
        flyout.style.width = `${width}px`;
        flyout.style.setProperty('--mgt-flyout-set-width', `${width}px`);
        window.requestAnimationFrame(() => this.updateFlyout());
      }
    }
  }

  private setupWindowEvents(isOpen: boolean) {
    if (isOpen && this.isLightDismiss) {
      window.addEventListener('wheel', this.handleWindowEvent);
      window.addEventListener('pointerdown', this.handleWindowEvent);
      window.addEventListener('resize', this.handleResize);
      window.addEventListener('keyup', this.handleKeyUp);
    } else {
      window.removeEventListener('wheel', this.handleWindowEvent);
      window.removeEventListener('pointerdown', this.handleWindowEvent);
      window.removeEventListener('resize', this.handleResize);
      window.removeEventListener('keyup', this.handleKeyUp);
    }
  }

  private handleWindowEvent(e: Event) {
    const flyout = this._flyout;

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

    this.close();
  }

  private handleResize(e: Event) {
    this.close();
  }

  private handleKeyUp(e: KeyboardEvent) {
    if (e.key === 'Escape') {
      this.close();
    }
  }

  private handleFlyoutWheel(e: Event) {
    e.preventDefault();
  }
}
