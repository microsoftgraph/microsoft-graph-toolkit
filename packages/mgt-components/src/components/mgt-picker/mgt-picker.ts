/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { MgtTemplatedComponent, mgtHtml, customElement } from '@microsoft/mgt-element';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import '../../styles/style-helper';
import { Entity } from '@microsoft/microsoft-graph-types';
import { DataChangedDetail } from '../mgt-get/mgt-get';

registerFluentComponents(fluentCombobox, fluentOption);

/**
 * Web component that allows a single entity pick from a generic endpoint from Graph. Uses mgt-get.
 *
 * @fires {CustomEvent<any>} selectionChanged - Fired when an option is clicked/selected
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --picker-background-color - {Color} Picker component background color
 */
@customElement('picker')
export class MgtPicker extends MgtTemplatedComponent {
  protected get strings() {
    return strings;
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

  private isRefreshing: boolean;

  @state() private response: Entity[];

  constructor() {
    super();
    this.placeholder = this.strings.comboboxPlaceholder;
    this.entityType = null;
    this.keyName = null;
    this.isRefreshing = false;
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
    this.isRefreshing = true;
    if (hardRefresh) {
      this.clearState();
    }
    void this.requestStateUpdate(hardRefresh);
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

  /**
   * Invoked on each update to perform rendering the picker. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    if (this.isLoadingState && !this.response) {
      return this.renderTemplate('loading', null);
    } else if (this.hasTemplate('error')) {
      const error = this.error ? (this.error as Error) : null;
      return this.renderTemplate('error', { error }, 'error');
    } else if (this.hasTemplate('no-data')) {
      return this.renderTemplate('no-data', null);
    }

    return this.response?.length > 0 ? this.renderPicker() : this.renderGet();
  }

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
          <fluent-option
            value=${item.id}
            @click=${(e: MouseEvent) => this.handleClick(e, item)}>
              ${item[this.keyName]}
          </fluent-option>`
        )}
      </fluent-combobox>
     `;
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
        resource=${this.resource}
        version=${this.version}
        .scopes=${this.scopes}
        max-pages=${this.maxPages}
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`;
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtPicker
   */
  protected async loadState() {
    if (!this.response) {
      const parent = this.renderRoot.querySelector('mgt-get');
      parent.addEventListener('dataChange', (e: CustomEvent<DataChangedDetail>): void => this.handleDataChange(e));
    }
    this.isRefreshing = false;
    // hack to maintain method signature contract
    await Promise.resolve();
  }

  private handleDataChange(e: CustomEvent<DataChangedDetail>): void {
    const response = e.detail.response.value;
    const error = e.detail.error ? e.detail.error : null;
    this.response = response;
    this.error = error;
  }

  private handleClick(e: MouseEvent, item: any) {
    this.fireCustomEvent('selectionChanged', item, true, false, true);
  }

  /**
   * Handles getting the fluent option item in the dropdown and fires a custom
   * event with it when you press Enter or Backspace keys.
   *
   * @param {KeyboardEvent} e
   */
  private handleComboboxKeydown = (e: KeyboardEvent) => {
    let value: string;
    let item: any;
    const keyName: string = e.key;
    const comboBox: HTMLElement = e.target as HTMLElement;
    const fluentOptionEl = comboBox.querySelector('.selected');
    if (fluentOptionEl) {
      value = fluentOptionEl.getAttribute('value');
    }

    if (['Enter'].includes(keyName)) {
      if (value) {
        item = this.response.filter(res => res.id === value).pop();
        this.fireCustomEvent('selectionChanged', item, true, false, true);
      }
    }
  };
}
