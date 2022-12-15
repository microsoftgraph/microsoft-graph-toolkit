/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { customElement, property } from 'lit/decorators.js';
import { MgtTemplatedComponent } from '@microsoft/mgt-element';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import '../../styles/style-helper';

registerFluentComponents(fluentCombobox, fluentOption);

/**
 * Web component that allows a single entity pick from a generic endpoint from Graph. Uses mgt-get.
 z
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-picker')
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
    reflect: true,
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
    reflect: true,
    type: String
  })
  public version: string = 'v1.0';

  /**
   * Maximum number of pages to get for the resource
   * default = 3
   * if <= 0, all pages will be fetched
   *
   * @type {number}
   * @memberof MgtGet
   */
  @property({
    attribute: 'max-pages',
    reflect: true,
    type: Number
  })
  public maxPages: number = 3;

  /**
   * Set the person details to render
   *
   * @type {object}
   * @memberof MgtPicker
   */
  @property({
    attribute: 'response',
    type: Object
  })
  public get response(): any {
    return this._response;
  }
  public set response(value) {
    if (this._response === value) {
      return;
    }

    this._response = value;
    this.requestStateUpdate(true);
  }

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
  public get placeholder(): string {
    return this._placeholder;
  }
  public set placeholder(value) {
    if (this._placeholder === value) {
      return;
    }
    this._placeholder = value;
    this.requestStateUpdate(true);
  }

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
    },
    reflect: true
  })
  public scopes: string[] = [];

  private _placeholder: string;
  private _response: object;

  constructor() {
    super();
    this._placeholder = 'Select an item';
  }

  /**
   * Clears the state of the component
   *
   * @protected
   * @memberof MgtPicker
   */
  protected clearState(): void {
    // this.response = null;
  }

  /**
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected requestStateUpdate(force?: boolean) {
    if (force) {
      //   this.people = null;
    }
    return super.requestStateUpdate(force);
  }

  /**
   * Invoked on each update to perform rendering the picker. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    if (this.isLoadingState) {
      return this.renderLoading();
    }
    console.log('response:', this.response);

    return this.renderPicker(this.resource, this.scopes);
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtPicker
   */
  protected renderLoading() {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render picker.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPicker
   */
  protected renderPicker(resource: string, scopes: string[]): TemplateResult {
    return html`
    <mgt-get resource=${resource} version=${this.version} scopes=${scopes}>
      <template>
        <fluent-combobox id="combobox" autocomplete="list">
          <fluent-option value={{item.id}} data-for="item in value"> {{ item.displayName }} </fluent-option>
        </fluent-combobox>
      </template>
    </mgt-get>
     `;
  }

  /**
   * render the no data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPicker
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
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
      parent.addEventListener('dataChange', e => this.handleDataChange(e));
    }
  }

  private handleDataChange(e) {
    console.log('here');
    let response = e.detail.response.value;
    this.response = response;
  }
}
