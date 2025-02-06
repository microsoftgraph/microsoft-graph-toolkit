/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Task } from '@lit/task';

import { LitElement, PropertyValueMap, PropertyValues, TemplateResult, html } from 'lit';
import { state } from 'lit/decorators.js';
import { ProviderState } from '../providers/IProvider';
import { Providers } from '../providers/Providers';
import { LocalizationHelper } from '../utils/LocalizationHelper';
import { PACKAGE_VERSION } from '../utils/version';
import { ComponentMediaQuery } from './baseComponent';

/**
 * BaseComponent extends LitElement adding mgt specific features to all components
 *
 * @export  MgtBaseTaskComponent
 * @abstract
 * @class MgtBaseTaskComponent
 * @extends {LitElement}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 */
export abstract class MgtBaseTaskComponent extends LitElement {
  /**
   * Supplies the component with a reactive property based on the current provider state
   *
   * @protected
   * @memberof MgtBaseTaskComponent
   */
  @state() protected providerState: ProviderState = ProviderState.Loading;
  /**
   * Exposes the semver of the library the component is part of
   *
   * @readonly
   * @static
   * @memberof MgtBaseTaskComponent
   */
  public static get packageVersion() {
    return PACKAGE_VERSION;
  }

  /**
   * Gets or sets the direction of the component
   *
   * @protected
   * @memberof MgtBaseTaskComponent
   */
  @state() protected direction: 'ltr' | 'rtl' | 'auto' = 'ltr';

  /**
   * Gets the ComponentMediaQuery of the component
   *
   * @readonly
   * @type {MgtElement.ComponentMediaQuery}
   * @memberof MgtBaseTaskComponent
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
   * A flag to check if the component has updated once.
   *
   * @readonly
   * @protected
   * @type {boolean}
   * @memberof MgtBaseTaskComponent
   */
  protected get isFirstUpdated(): boolean {
    return this._isFirstUpdated;
  }

  /**
   * returns component strings
   *
   * @readonly
   * @protected
   * @memberof MgtBaseTaskComponent
   */
  protected get strings(): Record<string, string> {
    return {};
  }

  private _isFirstUpdated = false;

  constructor() {
    super();
    this.handleDirectionChanged();
    this.handleLocalizationChanged();
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtBaseTaskComponent
   */
  public connectedCallback() {
    super.connectedCallback();
    LocalizationHelper.onStringsUpdated(this.handleLocalizationChanged);
    LocalizationHelper.onDirectionUpdated(this.handleDirectionChanged);
  }

  /**
   * Invoked each time the custom element is removed from a document-connected element
   *
   * @memberof MgtBaseTaskComponent
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
  protected firstUpdated(changedProperties: PropertyValueMap<unknown> | Map<PropertyKey, unknown>): void {
    super.firstUpdated(changedProperties);
    this._isFirstUpdated = true;
    Providers.onProviderUpdated(this.handleProviderUpdates);
    Providers.onActiveAccountChanged(this.handleActiveAccountUpdates);
  }

  /**
   * Used to clear state in inherited components
   */
  protected clearState(): void {
    // no-op
  }

  /**
   * helps facilitate creation of events across components
   *
   * @protected
   * @param {string} eventName
   * @param {*} [detail]
   * @param {boolean} [bubbles=false]
   * @param {boolean} [cancelable=false]
   * @param {boolean} [composed=false]
   * @return {*}  {boolean}
   * @memberof MgtBaseTaskComponent
   */
  protected fireCustomEvent(
    eventName: string,
    detail?: unknown,
    bubbles = false,
    cancelable = false,
    composed = false
  ): boolean {
    const event = new CustomEvent(eventName, {
      bubbles,
      cancelable,
      composed,
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
    this.fireCustomEvent('updated', undefined, true, false);
  }

  /**
   * load state into the component.
   * Override this function to provide actual loading logic.
   */
  protected async loadState() {
    return Promise.resolve();
  }

  /**
   * Override this function to provide the actual list of properties to trigger the task to run.
   * The default implementation returns an array with the providerState.
   * @returns {unknown[]} the properties when changed which trigger the Task to run
   */
  protected args(): unknown[] {
    return [this.providerState];
  }

  /**
   * Task that is run whenever one of the args changes
   * By default this task will call loadState
   */
  protected _task = new Task(this, {
    task: () => this.loadState(),
    args: () => this.args()
  });

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return this._task.render({
      pending: this.renderLoading,
      complete: this.renderContent,
      error: this.renderError
    });
  }

  /**
   * A default loading template.
   * @returns default loading template
   */
  protected renderLoading = (): TemplateResult => {
    return html`<span>Loading...</span>`;
  };

  protected renderError = (e: unknown): TemplateResult => {
    return html`<p>Error: ${e}</p>`;
  };

  protected renderContent = (): TemplateResult => {
    return html`<!-- baseTaskComponent, please implement renderContent -->`;
  };

  private readonly handleProviderUpdates = () => {
    this.providerState = Providers.globalProvider?.state ?? ProviderState.Loading;
  };

  private readonly handleActiveAccountUpdates = () => {
    this.clearState();
  };

  private readonly handleLocalizationChanged = () => {
    LocalizationHelper.updateStringsForTag(this.tagName, this.strings);
    this.requestUpdate();
  };

  private readonly handleDirectionChanged = () => {
    this.direction = LocalizationHelper.getDocumentDirection();
  };
}
