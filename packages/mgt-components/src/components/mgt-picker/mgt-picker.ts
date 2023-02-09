/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { MgtTemplatedComponent, mgtHtml, customElement } from '@microsoft/mgt-element';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import '../../styles/style-helper';

registerFluentComponents(fluentCombobox, fluentOption);

/**
 * Web component that allows a single entity pick from a generic endpoint from Graph. Uses mgt-get.
 *
 * @fires {CustomEvent<any>} selectionChanged - Fired when an option is clicked/selected
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedComponent}
 */
// @customElement('mgt-picker')
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
  public version: string = 'v1.0';

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
  public maxPages: number = 3;

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
  public cacheEnabled: boolean = false;

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
  public cacheInvalidationPeriod: number = 0;

  private isRefreshing: boolean;

  @state() private response: any[];
  @state() private error: any[];

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
    this.requestStateUpdate(hardRefresh);
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
      return this.renderTemplate('error', this.error ? this.error : null, 'error');
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
      <fluent-combobox id="combobox" autocomplete="list" placeholder=${this.placeholder}>
        ${this.response.map(
          item => html`
          <fluent-option value=${item.id} @click=${e => this.handleClick(e, item)}> ${
            item[this.keyName]
          } </fluent-option>`
        )}
      </fluent-combobox>
     `;
  }

  /**
   * Render picker.
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
        scopes=${this.scopes} 
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
      let parent = this.renderRoot.querySelector('mgt-get');
      parent.addEventListener('dataChange', (e): void => this.handleDataChange(e));
    }
    this.isRefreshing = false;
  }

  private handleDataChange(e: any): void {
    let response = e.detail.response.value;
    let error = e.detail.error ? e.detail.error : null;
    this.response = response;
    this.error = error;
  }

  private handleClick(e: MouseEvent, item: any) {
    this.fireCustomEvent('selectionChanged', item);
  }
}
