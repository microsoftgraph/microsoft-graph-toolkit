/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { customElement, property } from 'lit/decorators.js';
import { IGraph } from '@microsoft/mgt-element';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import { styles } from './mgt-picker-css';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { getTodoTaskLists, TodoTaskList } from '../mgt-todo/graph.todo';
import '../../styles/style-helper';

registerFluentComponents(fluentCombobox, fluentOption);

/**
 * Web component that allows a single entity pick from a generic endpoint from Graph. Uses mgt-get.
 
 * @export
 * @class MgtGenericPicker
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-picker')
export class MgtGenericPicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
  }

  /**
   * The resource to get
   *
   * @type {string}
   * @memberof MgtGenericPicker
   */
  @property({
    attribute: 'resource',
    type: String
  })
  public get resource(): string {
    return this._resource;
  }
  public set resource(value) {
    if (this._resource === value) {
      return;
    }
    this._resource = value;
    this.requestStateUpdate(true);
  }

  /**
   * Api version to use for request
   *
   * @type {string}
   * @memberof MgtGenericPicker
   */
  @property({
    attribute: 'version',
    type: String
  })
  public get version(): string {
    return this._version;
  }
  public set version(value) {
    if (this._version === value) {
      return;
    }
    this._version = value;
    this.requestStateUpdate(true);
  }

  /**
   * The scopes to request
   *
   * @type {string[]}
   * @memberof MgtGenericPicker
   */
  @property({
    attribute: 'scopes',
    converter: value => {
      return value ? value.toLowerCase().split(',') : null;
    },
    reflect: true
  })
  public scopes: string[] = [];

  private _resource: string;
  private _version: string = 'v1.0';
  private _lists: TodoTaskList[];
  private _graph: IGraph;

  constructor() {
    super();
    this._lists = [];
    this._graph = null;
  }

  /**
   * Clears the state of the component
   *
   * @protected
   * @memberof MgtGenericPicker
   */
  protected clearState(): void {
    // this.people = null;
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
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    if (this.isLoadingState) {
      return this.renderLoading();
    }
    return this.renderTemplate('default', { entity: this._lists }) || this.renderPicker(this.resource, this.scopes);
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtGenericPicker
   */
  protected renderLoading() {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render picker.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtGenericPicker
   */
  protected renderPicker(resource: String, scopes: String[]): TemplateResult {
    return html`
      <mgt-get id="entityGet" resource=${resource} version=${this.version} scopes=${scopes}>
        <template>
          <fluent-combobox id="combobox" placeholder="Select a task list" autocomplete="both">
            <div data-for="item in value">
              <fluent-option value={{item.id}}> {{ item.displayName }} </fluent-option>
            </div>
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
   * @memberof MgtGenericPicker
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtGenericPicker
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedOut) {
      return;
    }

    if (!this._graph) {
      const graph = provider.graph.forComponent(this);
      this._graph = graph;
    }

    let lists = this._lists;
    if (!lists || !lists.length) {
      lists = await getTodoTaskLists(this._graph);
      this._lists = lists;
    }
    console.log('Lists: ', this._lists);
  }
}
