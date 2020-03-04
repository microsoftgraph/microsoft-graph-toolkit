/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { LitElement, property, PropertyValues } from 'lit-element';
import { Providers } from '../Providers';

/**
 * Defines media query based on component width
 *
 * @export
 * @enum {string}
 */
export enum ComponentMediaQuery {
  /**
   * devices with width < 768
   */
  mobile = '',

  /**
   * devies with width < 1200
   */
  tablet = 'tablet',

  /**
   * devices with width > 1200
   */
  desktop = 'desktop'
}

/**
 * BaseComponent extends LitElement including ShadowRoot toggle and fireCustomEvent features
 *
 * @export  MgtBaseComponent
 * @abstract
 * @class MgtBaseComponent
 * @extends {LitElement}
 */
export abstract class MgtBaseComponent extends LitElement {
  /**
   * Get ShadowRoot toggle, returns value of _useShadowRoot
   *
   * @static _useShadowRoot
   * @memberof MgtBaseComponent
   */
  public static get useShadowRoot() {
    return this._useShadowRoot;
  }

  /**
   * Set ShadowRoot toggle value
   *
   * @static _useShadowRoot
   * @memberof MgtBaseComponent
   */
  public static set useShadowRoot(value: boolean) {
    this._useShadowRoot = value;
  }

  /**
   * Gets the ComponentMediaQuery of the component
   *
   * @readonly
   * @type {ComponentMediaQuery}
   * @memberof MgtBaseComponent
   */
  public get mediaQuery(): ComponentMediaQuery {
    if (this.offsetWidth < 768) {
      return ComponentMediaQuery.mobile;
    } else if (this.offsetWidth < 1200) {
      return ComponentMediaQuery.tablet;
    } else {
      return ComponentMediaQuery.desktop;
    }
  }

  /**
   * A flag to check if the component's firstUpdated method has fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected get isLoadingState(): boolean {
    return this._isLoadingState;
  }

  private static _useShadowRoot: boolean = true;

  /**
   * determines if login component is in loading state
   * @type {boolean}
   */
  @property({ attribute: false })
  private _isLoadingState: boolean = false;

  private _cts: CancellationTokenSource;

  constructor() {
    super();
    if (this.isShadowRootDisabled()) {
      (this as any)._needsShimAdoptedStyleSheets = true;
    }
  }

  /**
   * Receive ShadowRoot Disabled value
   *
   * @returns boolean _useShadowRoot value
   * @memberof MgtBaseComponent
   */
  public isShadowRootDisabled() {
    return !MgtBaseComponent._useShadowRoot || !(this.constructor as typeof MgtBaseComponent)._useShadowRoot;
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * @param _changedProperties Map of changed properties with old values
   */
  protected firstUpdated(changedProperties): void {
    super.firstUpdated(changedProperties);
    Providers.onProviderUpdated(() => this.reloadState());
    this.reloadState();
  }

  /**
   * load state into the component.
   * Override this function to provide additional loading logic.
   */
  protected loadState(cancellationToken?: CancellationToken): Promise<void> {
    return Promise.resolve();
  }

  /**
   * helps facilitate creation of events across components
   *
   * @protected
   * @param {string} eventName name given to specific event
   * @param {*} [detail] optional any value to dispatch with event
   * @returns {boolean}
   * @memberof MgtBaseComponent
   */
  protected fireCustomEvent(eventName: string, detail?: any): boolean {
    const event = new CustomEvent(eventName, {
      bubbles: false,
      cancelable: true,
      detail
    });
    return this.dispatchEvent(event);
  }

  /**
   * method to create ShadowRoot if disabled flag isn't present
   *
   * @protected
   * @returns boolean
   * @memberof MgtBaseComponent
   */
  protected createRenderRoot() {
    return this.isShadowRootDisabled() ? this : super.createRenderRoot();
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
  protected updated(changedProperties: PropertyValues) {
    super.updated(changedProperties);
    const event = new CustomEvent('updated', {
      bubbles: true,
      cancelable: true
    });
    this.dispatchEvent(event);
  }

  /**
   * Reload the state.
   * Returns false if already loading, unless forced.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtBaseComponent
   */
  protected async reloadState(force: boolean = false): Promise<boolean> {
    if (!force && this._isLoadingState) {
      return false;
    }

    if (this._cts) {
      this._cts.cancel();
    }
    this._cts = new CancellationTokenSource();

    try {
      this._isLoadingState = true;
      this.fireCustomEvent('loadingInitiated');

      await this.loadState(this._cts.token);
      this.fireCustomEvent('loadingCompleted');
    } catch (e) {
      if (e instanceof OperationCancelledException) {
        this.fireCustomEvent('loadingCancelled');
        return false;
      } else {
        throw e;
      }
    } finally {
      this._isLoadingState = false;
    }

    return true;
  }
}

/**
 * Manages a CancellationToken and provides a mechanism for cancelling a function.
 *
 * @export
 * @class CancellationTokenSource
 */
// tslint:disable-next-line: max-classes-per-file
export class CancellationTokenSource {
  private _token: CancellationToken;

  /**
   * is the token in a cancelled state?
   *
   * @readonly
   * @type {boolean}
   * @memberof CancellationTokenSource
   */
  public get isCancellationRequested(): boolean {
    return this._token.isCancellationRequested;
  }

  /**
   * Lazily retrieve a CancellationToken instance.
   *
   * @readonly
   * @type {CancellationToken}
   * @memberof CancellationTokenSource
   */
  public get token(): CancellationToken {
    if (!this._token) {
      this._token = new CancellationToken();
    }
    return this._token;
  }

  /**
   * set the token to a cancelled state.
   *
   * @memberof CancellationTokenSource
   */
  public cancel(): void {
    this._token.requestCancellation();
  }
}

/**
 * A CancellationToken is used in conjunction with a CancellationTokenSource to support a method
 * of interupting process flow upon request.
 *
 * @export
 * @class CancellationToken
 */
// tslint:disable-next-line: max-classes-per-file
export class CancellationToken {
  private _isCancellationRequested: boolean;

  /**
   * is the token in a cancelled state?
   *
   * @readonly
   * @type {boolean}
   * @memberof CancellationToken
   */
  public get isCancellationRequested(): boolean {
    return this._isCancellationRequested;
  }

  /**
   * sets the token to a cancelled state.
   *
   * @memberof CancellationToken
   */
  public requestCancellation(): void {
    this._isCancellationRequested = true;
  }

  /**
   * interupts the current process if cancellation has been requested.
   *
   * @memberof CancellationToken
   */
  public throwIfCancellationRequested(): void {
    if (this._isCancellationRequested) {
      throw new OperationCancelledException();
    }
  }
}

/**
 * A dummy type used to detect when a thrown exception is caused by a token cancellation.
 *
 * @export
 * @class OperationCancelledException
 */
// tslint:disable-next-line: max-classes-per-file
export class OperationCancelledException {}
