/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TermStore } from '@microsoft/microsoft-graph-types';
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
 * @fires {CustomEvent<TermStore.Term>} selectionChanged - Fired when an option is clicked/selected
 * @export
 * @class MgtTaxonomyPicker
 * @extends {MgtTemplatedComponent}
 */
// @customElement('mgt-picker')
@customElement('taxonomy-picker')
export class MgtTaxonomyPicker extends MgtTemplatedComponent {
  protected get strings() {
    return strings;
  }

  /**
   * The termSetId of the term set whose children to get.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'termset-id',
    type: String
  })
  public termSetId: string;

  /**
   * The termId of the term whose children to get. This term must be a child of the term set specified by termSetId.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'term-id',
    type: String
  })
  public termId: string;

  /**
   * The locale based on which the term names should be returned.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'locale',
    type: String
  })
  public locale: string;

  /**
   * Api version to use for request.
   * Default is beta.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'version',
    type: String
  })
  public version: string = 'beta';

  /**
   * A placeholder for the picker.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * The default selected term id.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'default-selected-term-id',
    type: String
  })
  public get defaultSelectedTermId(): string {
    return this._defaultSelectedTermId;
  }
  public set defaultSelectedTermId(value: string) {
    this._defaultSelectedTermId = value;
  }

  /**
   * The selected term.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'selected-term',
    type: Object
  })
  public get selectedTerm(): TermStore.Term {
    return this._selectedTerm;
  }
  public set selectedTerm(value: TermStore.Term) {
    this._selectedTerm = value;
  }

  /**
   * Enables cache on the response from the specified resource.
   * Default is false.
   *
   * @type {boolean}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'cache-enabled',
    type: Boolean
  })
  public cacheEnabled: boolean = false;

  /**
   * Invalidation period of the cache for the responses in milliseconds.
   *
   * @type {number}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'cache-invalidation-period',
    type: Number
  })
  public cacheInvalidationPeriod: number = 0;

  private isRefreshing: boolean;
  private _selectedTerm: TermStore.Term;
  private _defaultSelectedTermId: string;

  @state() private terms: TermStore.Term[];
  @state() private noTerms: boolean;
  @state() private error: any;

  constructor() {
    super();
    this.placeholder = this.strings.comboboxPlaceholder;
    this.isRefreshing = false;
    this.noTerms = false;
  }

  /**
   * Refresh the data
   *
   * @param {boolean} [hardRefresh=false]
   * if false (default), the component will only update if the data changed
   * if true, the data will be first cleared and reloaded completely
   * @memberof MgtTaxonomyPicker
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
   * @memberof MgtTaxonomyPicker
   */
  protected clearState(): void {
    this.terms = null;
    this.error = null;
    this.noTerms = false;
  }

  /**
   * Invoked on each update to perform rendering the picker. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public render() {
    // if loading state, render loading template
    if (this.isLoadingState && !this.terms) {
      return this.renderLoading();
    }

    // if error state, render error template
    if (this.error) {
      return this.renderError();
    }

    // if no terms are found, render no-data template
    if (this.noTerms) {
      return this.renderNoData();
    }

    // if terms are found, render picker else render get
    return this.terms?.length > 0 ? this.renderTaxonomyPicker() : this.renderGet();
  }

  /**
   * Renders loading spinner while terms are fetched from the Graph
   *
   * @protected
   * @returns
   * @memberof MgtTaxonomyPicker
   */
  protected renderLoading(): TemplateResult {
    return (
      this.renderTemplate('loading', null, 'loading') ||
      mgtHtml`
        <div class="message-parent">
          <mgt-spinner></mgt-spinner>
          <div label="loading-text" aria-label="loading">
            ${this.strings.loadingMessage}
          </div>
        </div>
      `
    );
  }

  /**
   * Render the no-data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderError(): TemplateResult {
    return (
      this.renderTemplate('error', null, 'error') ||
      html`
              <span>
                ${this.error.message}
            </span>
          `
    );
  }
  /**
   * Render the no-data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderNoData(): TemplateResult {
    return (
      this.renderTemplate('no-data', null) ||
      html`
            <span>
              ${this.strings.noTermsFound}
            </span>
          `
    );
  }

  /**
   * Render picker.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderTaxonomyPicker(): TemplateResult {
    return mgtHtml`
      <fluent-combobox id="taxonomy-picker" autocomplete="list" placeholder=${this.placeholder}>
        ${this.terms.map(term => this.renderTaxonomyPickerItem(term))}
      </fluent-combobox>
     `;
  }

  /**
   * Render picker item.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderTaxonomyPickerItem(term: TermStore.Term): TemplateResult {
    return this.defaultSelectedTermId && this.defaultSelectedTermId === term.id
      ? html`
        <fluent-option value=${term.id} selected @click=${e => this.handleClick(e, term)}> ${
          this.renderTemplate('term', { term }, term.id) || term.labels[0].name
        } </fluent-option>
        `
      : html`
        <fluent-option value=${term.id} @click=${e => this.handleClick(e, term)}> ${
          this.renderTemplate('term', { term }, term.id) || term.labels[0].name
        } </fluent-option>
        `;
  }

  /**
   * Render picker.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderGet(): TemplateResult {
    // if termSetId is not specified, return error message
    if (!this.termSetId) {
      return html`
            <span>
                ${this.strings.termSetIdRequired}
            </span>
            `;
    }

    let resource: string = `/termStore/sets/${this.termSetId}/children`;

    // if both termSetId and termId are specified, then set resource to /termStore/sets/{termSetId}/terms/{termId}/children
    if (this.termId) {
      resource = `/termStore/sets/${this.termSetId}/terms/${this.termId}/children`;
    }

    return mgtHtml`
      <mgt-get 
        resource=${resource}
        version=${this.version} 
        scopes=${['TermStore.Read.All']}  
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`;
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtTaxonomyPicker
   */
  protected async loadState() {
    if (!this.terms) {
      let parent = this.renderRoot.querySelector('mgt-get');
      parent.addEventListener('dataChange', (e): void => this.handleDataChange(e));
    }
    this.isRefreshing = false;
  }

  private handleDataChange(e: any): void {
    let error = e.detail.error ? e.detail.error : null;

    if (error) {
      this.error = error;
      return;
    }

    let response = e.detail.response.value;

    // if response is not null and has values, if locale is specified, then
    // get the label in response that has languageTag equal to locale and make it the first label and append the rest of the labels

    let terms = response.map((item: TermStore.Term) => {
      let labels = item.labels;
      if (labels && labels.length > 0) {
        if (this.locale) {
          let label = labels.find(label => label.languageTag.toLowerCase() === this.locale.toLowerCase());
          if (label) {
            item.labels = [label, ...labels.filter(label => label.languageTag !== this.locale)];
          }
        }
      }
      return item;
    });

    this.terms = terms;

    //  if there are no terms then set noTerms to true
    if (terms.length === 0) {
      this.noTerms = true;
    }
  }

  private handleClick(e: MouseEvent, item: TermStore.Term) {
    this.selectedTerm = item;
    this.fireCustomEvent('selectionChanged', item);
  }
}
