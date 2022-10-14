/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { internalProperty, LitElement, PropertyValues } from 'lit-element';
import { ProviderState } from '../providers/IProvider';
import { Providers } from '../providers/Providers';
import { LocalizationHelper } from '../utils/LocalizationHelper';

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
   * devices with width < 1200
   */
  tablet = 'tablet',

  /**
   * devices with width > 1200
   */
  desktop = 'desktop'
}

/**
 * BaseComponent extends LitElement adding mgt specific features to all components
 *
 * @export  MgtBaseComponent
 * @abstract
 * @class MgtBaseComponent
 * @extends {LitElement}
 */
export abstract class MgtBaseComponent extends LitElement {
  /**
   * Gets or sets the direction of the component
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  @internalProperty() protected direction = 'ltr';

  /**
   * Gets the ComponentMediaQuery of the component
   *
   * @readonly
   * @type {MgtElement.ComponentMediaQuery}
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
   * A flag to check if the component is loading data state.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected get isLoadingState(): boolean {
    return this._isLoadingState;
  }

  /**
   * A flag to check if the component has updated once.
   *
   * @readonly
   * @protected
   * @type {boolean}
   * @memberof MgtBaseComponent
   */
  protected get isFirstUpdated(): boolean {
    return this._isFirstUpdated;
  }

  /**
   * returns component strings
   *
   * @readonly
   * @protected
   * @memberof MgtBaseComponent
   */
  protected get strings(): { [x: string]: string } {
    return {};
  }

  /**
   * determines if login component is in loading state
   * @type {boolean}
   */
  private _isLoadingState: boolean = false;

  private _isFirstUpdated = false;
  private _currentLoadStatePromise: Promise<unknown>;

  constructor() {
    super();
    this.handleLocalizationChanged = this.handleLocalizationChanged.bind(this);
    this.handleDirectionChanged = this.handleDirectionChanged.bind(this);
    this.handleProviderUpdates = this.handleProviderUpdates.bind(this);
    this.handleActiveAccountUpdates = this.handleActiveAccountUpdates.bind(this);
    this.handleDirectionChanged();
    this.handleLocalizationChanged();
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtBaseComponent
   */
  public connectedCallback() {
    super.connectedCallback();
    LocalizationHelper.onStringsUpdated(this.handleLocalizationChanged);
    LocalizationHelper.onDirectionUpdated(this.handleDirectionChanged);
  }

  /**
   * Invoked each time the custom element is removed from a document-connected element
   *
   * @memberof MgtBaseComponent
   */
  public disconnectedCallback() {
    super.disconnectedCallback();
    LocalizationHelper.removeOnStringsUpdated(this.handleLocalizationChanged);
    LocalizationHelper.removeOnDirectionUpdated(this.handleDirectionChanged);
    Providers.removeProviderUpdatedListener(this.handleProviderUpdates);
    Providers.removeActiveAccountChangedListener(this.handleActiveAccountUpdates);
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
    this._isFirstUpdated = true;
    Providers.onProviderUpdated(this.handleProviderUpdates);
    Providers.onActiveAccountChanged(this.handleActiveAccountUpdates);
    this.requestStateUpdate();
  }

  /**
   * load state into the component.
   * Override this function to provide additional loading logic.
   */
  protected loadState(): Promise<void> {
    return Promise.resolve();
  }

  protected clearState(): void {}

  /**
   * helps facilitate creation of events across components
   *
   * @protected
   * @param {string} eventName
   * @param {*} [detail]
   * @param {boolean} [bubbles=false]
   * @param {boolean} [cancelable=false]
   * @return {*}  {boolean}
   * @memberof MgtBaseComponent
   */
  protected fireCustomEvent(
    eventName: string,
    detail?: any,
    bubbles: boolean = false,
    cancelable: boolean = false
  ): boolean {
    const event = new CustomEvent(eventName, {
      bubbles,
      cancelable,
      detail
    });
    return this.dispatchEvent(event);
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
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected async requestStateUpdate(force: boolean = false): Promise<unknown> {
    // the component is still bootstraping - wait until first updated
    if (!this._isFirstUpdated) {
      return;
    }

    // Wait for the current load promise to complete (unless forced).
    if (this.isLoadingState && !force) {
      await this._currentLoadStatePromise;
    }

    const provider = Providers.globalProvider;

    if (!provider) {
      return Promise.resolve();
    }

    if (provider.state === ProviderState.SignedOut) {
      // Signed out, clear the component state
      this.clearState();
      return;
    } else if (provider.state === ProviderState.Loading) {
      // The provider state is indeterminate. Do nothing.
      return Promise.resolve();
    } else {
      // Signed in, load the internal component state
      const loadStatePromise = new Promise<void>(async (resolve, reject) => {
        try {
          this.setLoadingState(true);
          this.fireCustomEvent('loadingInitiated');

          await this.loadState();

          this.setLoadingState(false);
          this.fireCustomEvent('loadingCompleted');
          resolve();
        } catch (e) {
          // Loading failed. Clear any partially set data.
          this.clearState();

          this.setLoadingState(false);
          this.fireCustomEvent('loadingFailed');
          reject(e);
        }

        // Return the load state promise.
        // If loading + forced, chain the promises.
        // This is to account for the lack of a cancellation token concept.
        return (this._currentLoadStatePromise =
          this.isLoadingState && !!this._currentLoadStatePromise && force
            ? this._currentLoadStatePromise.then(() => loadStatePromise)
            : loadStatePromise);
      });
    }
  }

  private setLoadingState(value: boolean) {
    if (this._isLoadingState === value) {
      return;
    }

    this._isLoadingState = value;
    this.requestUpdate('isLoadingState');
  }

  private handleProviderUpdates() {
    this.requestStateUpdate();
  }

  private async handleActiveAccountUpdates() {
    this.clearState();
    this.requestStateUpdate();
  }

  private handleLocalizationChanged() {
    LocalizationHelper.updateStringsForTag(this.tagName, this.strings);
    this.requestUpdate();
  }

  private handleDirectionChanged() {
    this.direction = LocalizationHelper.getDocumentDirection();
  }
}
