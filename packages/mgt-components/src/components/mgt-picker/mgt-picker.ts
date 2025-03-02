/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, PropertyValueMap, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { MgtTemplatedTaskComponent, mgtHtml } from '@microsoft/mgt-element';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import '../../styles/style-helper';
import { Entity } from '@microsoft/microsoft-graph-types';
import { DataChangedDetail, registerMgtGetComponent } from '../mgt-get/mgt-get';
import { styles } from './mgt-picker-css';
import { registerComponent } from '@microsoft/mgt-element';

export const registerMgtPickerComponent = () => {
  registerFluentComponents(fluentCombobox, fluentOption);

  registerMgtGetComponent();
  registerComponent('picker', MgtPicker);
};

/**
 * Web component that allows a single entity pick from a generic endpoint from Graph. Is a thin wrapper over mgt-get.
 * Does not load any state itself, only received state from mgt-get via events.
 *
 * @fires {CustomEvent<any>} selectionChanged - Fired when an option is clicked/selected
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedTaskComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 *
 * @cssprop --picker-background-color - {Color} Picker component background color
 * @cssprop --picker-list-max-height - {String} max height for options list. Default value is 380px.
 */
export class MgtPicker extends MgtTemplatedTaskComponent {
  protected get strings() {
    return strings;
  }

  public static get styles() {
    return styles;
  }

  /**
   * The resource to get
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'resource',
    type: String
  })
  public resource: string;

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'version',
    type: String
  })
  public version = 'v1.0';

  /**
   * Maximum number of pages to get for the resource
   * default = 3
   * if <= 0, all pages will be fetched
   *
   * @type {number}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'max-pages',
    type: Number
  })
  public maxPages = 3;

  /**
   * A placeholder for the picker
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Key to be rendered in the picker
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'key-name',
    type: String
  })
  public keyName: string;

  /**
   * Entity to be rendered in the picker
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'entity-type',
    type: String
  })
  public entityType: string;

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'scopes',
    converter: value => {
      return value ? value.toLowerCase().split(',') : null;
    }
  })
  public scopes: string[] = [];

  /**
   * Enables cache on the response from the specified resource
   * default = false
   *
   * @type {boolean}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'cache-enabled',
    type: Boolean
  })
  public cacheEnabled = false;

  /**
   * Invalidation period of the cache for the responses in milliseconds
   *
   * @type {number}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'cache-invalidation-period',
    type: Number
  })
  public cacheInvalidationPeriod = 0;

  /**
   * Sets the currently selected value for the picker
   * Must be present as an option in the supplied data returned from the the underlying graph query
   *
   * @type {string}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'selected-value',
    type: String
  })
  public selectedValue: string;

  @state() private response: Entity[];

  constructor() {
    super();
    this.placeholder = this.strings.comboboxPlaceholder;
    this.entityType = null;
    this.keyName = null;
  }

  /**
   * Refresh the data
   *
   * @param {boolean} [hardRefresh=false]
   * if false (default), the component will only update if the data changed
   * if true, the data will be first cleared and reloaded completely
   * @memberof MgtPicker
   */
  public refresh(hardRefresh = false) {
    if (hardRefresh) {
      this.clearState();
    }
    void this._task.run();
  }

  /**
   * Clears the state of the component
   *
   * @protected
   * @memberof MgtPicker
   */
  protected clearState(): void {
    this.response = null;
    this.error = null;
  }

  public renderLoading = (): TemplateResult => {
    if (!this.response) {
      return this.renderTemplate('loading', null);
    }
    return this.renderContent();
  };

  /**
   * Invoked on each update to perform rendering the picker. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public renderContent = () => {
    const error = this.error ? (this.error as Error) : null;
    if (error && this.hasTemplate('error')) {
      return this.renderTemplate('error', { error }, 'error');
    } else if (this.hasTemplate('no-data')) {
      return this.renderTemplate('no-data', null);
    }

    return this.response?.length > 0 ? this.renderPicker() : this.renderGet();
  };

  /**
   * Render picker.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPicker
   */
  protected renderPicker(): TemplateResult {
    return mgtHtml`
      <fluent-combobox
        @keydown=${this.handleComboboxKeydown}
        current-value=${ifDefined(this.selectedValue)}
        part="picker"
        class="picker"
        id="combobox"
        autocomplete="list"
        placeholder=${this.placeholder}>
          ${this.response.map(
            item => html`
            <fluent-option value=${item.id} @click=${(e: MouseEvent) =>
              this.handleClick(e, item)}> ${this.getNestedPropertyValue(item, this.keyName)} </fluent-option>`
          )}
      </fluent-combobox>
     `;
  }

  private getNestedPropertyValue(item: Entity, keyName: string) {
    const keys = keyName.split('.');
    let value: Entity | object | string = item;

    for (const key of keys) {
      value = value[key] as object | string;

      if (value === undefined) {
        console.warn(`mgt-picker: Key '${key}' is undefined.`);
        return '';
      }
    }

    return value;
  }

  /**
   * Renders mgt-get which does a GET request to the resource.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPicker
   */
  protected renderGet(): TemplateResult {
    return mgtHtml`
      <mgt-get
        class="mgt-get"
        resource=${this.resource}
        version=${this.version}
        .scopes=${this.scopes}
        max-pages=${this.maxPages}
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`;
  }

  /**
   * When the component is first updated wire up the event listeners.
   * @param changedProperties a map of changed properties with old values
   */
  protected firstUpdated(changedProperties: PropertyValueMap<unknown> | Map<PropertyKey, unknown>): void {
    super.firstUpdated(changedProperties);
    const parent = this.renderRoot;
    if (parent) {
      parent.addEventListener('dataChange', (e: CustomEvent<DataChangedDetail>): void => this.handleDataChange(e));
    } else {
      console.error('🦒: mgt-picker component requires a renderRoot. Something has gone horribly wrong.');
    }
  }

  private handleDataChange(e: CustomEvent<DataChangedDetail>): void {
    const response = e.detail.response.value;
    const error = e.detail.error ? e.detail.error : null;
    this.response = response;
    this.error = error;
  }

  private handleClick(_e: MouseEvent, item: Entity) {
    this.fireCustomEvent('selectionChanged', item, true, false, true);
  }

  /**
   * Handles getting the fluent option item in the dropdown and fires a custom
   * event with it when you press Enter or Backspace keys.
   *
   * @param {KeyboardEvent} e
   */
  private readonly handleComboboxKeydown = (e: KeyboardEvent) => {
    let value: string;
    let item: Entity;
    const keyName: string = e.key;
    const comboBox: HTMLElement = e.target as HTMLElement;
    const fluentOptionEl = comboBox.querySelector('.selected');
    if (fluentOptionEl) {
      value = fluentOptionEl.getAttribute('value');
    }

    if ('Enter' === keyName) {
      if (value) {
        item = this.response.filter(res => res.id === value).pop();
        this.fireCustomEvent('selectionChanged', item, true, false, true);
      }
    }
  };
}
