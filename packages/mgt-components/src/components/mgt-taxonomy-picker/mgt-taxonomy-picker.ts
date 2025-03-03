/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Position } from '../../graph/types';
import { html, PropertyValueMap, TemplateResult } from 'lit';
import { property, state } from 'lit/decorators.js';
import { MgtTemplatedTaskComponent, mgtHtml } from '@microsoft/mgt-element';
import { strings } from './strings';
import { fluentCombobox, fluentOption } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import '../../styles/style-helper';
import { styles } from './mgt-taxonomy-picker-css';
import { DataChangedDetail, registerMgtGetComponent } from '../mgt-get/mgt-get';
import { registerComponent } from '@microsoft/mgt-element';
import { registerMgtSpinnerComponent } from '../sub-components/mgt-spinner/mgt-spinner';

export const registerMgtTaxonomyPickerComponent = () => {
  registerFluentComponents(fluentCombobox, fluentOption);

  registerMgtSpinnerComponent();
  registerMgtGetComponent();
  registerComponent('taxonomy-picker', MgtTaxonomyPicker);
};

/**
 * Web component that can query the Microsoft Graph API for Taxonomy
 * and render a dropdown control with terms,
 * allowing selection of a single term based on
 * the specified term set id or a combination of the specified term set id and the specified term id.
 * Uses mgt-get.
 *
 * @fires {CustomEvent<MicrosoftGraph.TermStore.Term>} selectionChanged - Fired when an option is clicked/selected
 * @export
 * @class MgtTaxonomyPicker
 * @extends {MgtTemplatedTaskComponent}
 *
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
 *
 * @cssprop --taxonomy-picker-background-color - {Color} Picker component background color
 * @cssprop --taxonomy-picker-list-max-height - {String} max height for options list. Default value is 380px.
 * @cssprop --taxonomy-picker-placeholder-color - {Color} Text color for the placeholder in the picker
 * @cssprop --taxonomy-picker-placeholder-hover-color - {Color} Text color for the placeholder in the picker on hover
 */
export class MgtTaxonomyPicker extends MgtTemplatedTaskComponent {
  /**
   * The strings to be used for localizing the component.
   *
   * @readonly
   * @protected
   * @memberof MgtTaxonomyPicker
   */
  protected get strings() {
    return strings;
  }

  public static get styles() {
    return styles;
  }

  /**
   * The termsetId of the term set whose children to get.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'term-set-id',
    type: String
  })
  public termsetId: string;

  /**
   * The termId of the term whose children to get. This term must be a child of the term set specified by termsetId.
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
   * The id of the site where the termset is located. If not specified, the termset is assumed to be at the tenant level.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'site-id',
    type: String
  })
  public siteId: string;

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
  public version = 'beta';

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
   * The position of the dropdown. Can be 'above' or 'below'.
   *
   * @type {string}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'position',
    type: String,
    converter: (value: Position): Position => {
      if (value === 'above') {
        return 'above';
      }
      return 'below';
    }
  })
  public position: Position = 'below';

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
  public defaultSelectedTermId: string;

  /**
   * The selected term.
   *
   * @type {MicrosoftGraph.TermStore.Term}
   * @memberof MgtTaxonomyPicker
   */
  @property({
    attribute: 'selected-term',
    type: Object
  })
  public selectedTerm: MicrosoftGraph.TermStore.Term | null = null;

  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled: boolean;

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
  public cacheEnabled = false;

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
  public cacheInvalidationPeriod = 0;

  @state() private terms: MicrosoftGraph.TermStore.Term[];
  @state() private noTerms: boolean;
  // @state() private error: object;

  constructor() {
    super();
    this.placeholder = this.strings.comboboxPlaceholder;
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
    if (hardRefresh) {
      this.clearState();
    }
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
   * Renders loading spinner while terms are fetched from the Graph
   *
   * @protected
   * @returns
   * @memberof MgtTaxonomyPicker
   */
  protected renderLoading = () => {
    if (!this.terms) {
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
    return this.renderContent();
  };

  /**
   * Invoked on each update to perform rendering the picker. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  public renderContent = () => {
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
  };

  /**
   * Render the no-data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderError = (): TemplateResult =>
    this.renderTemplate('error', null, 'error') || html`<span>${this.error}</span>`;
  /**
   * Render the no-data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTaxonomyPicker
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html`<span>${this.strings.noTermsFound}</span>`;
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
      <fluent-combobox class="taxonomy-picker" autocomplete="both" placeholder=${this.placeholder} position=${
        this.position
      } ?disabled=${this.disabled}>
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
  protected renderTaxonomyPickerItem(term: MicrosoftGraph.TermStore.Term): TemplateResult {
    const selected: boolean = this.defaultSelectedTermId && this.defaultSelectedTermId === term.id;

    return html`
        <fluent-option value=${term.id} ?selected=${selected} @click=${(e: MouseEvent) => this.handleClick(e, term)}> ${
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
    // if termsetId is not specified, return error message
    if (!this.termsetId) {
      return html`<span>${this.strings.termsetIdRequired}</span>`;
    }

    let resource = `/termStore/sets/${this.termsetId}/children`;

    // if both termsetId and termId are specified, then set resource to /termStore/sets/{termsetId}/terms/{termId}/children
    if (this.termId) {
      resource = `/termStore/sets/${this.termsetId}/terms/${this.termId}/children`;
    }

    // if siteId is specified, then prefix /sites/{siteId}/ to the resource
    if (this.siteId) {
      resource = `/sites/${this.siteId}${resource}`;
    }

    // Add properties to select to the resource
    resource += '?$select=id,labels,descriptions,properties';

    return mgtHtml`
      <mgt-get
        class="mgt-get"
        resource=${resource}
        version=${this.version}
        scopes="TermStore.Read.All, TermStore.ReadWrite.All"
        ?cache-enabled=${this.cacheEnabled}
        ?cache-invalidation-period=${this.cacheInvalidationPeriod}>
      </mgt-get>`;
  }

  protected firstUpdated(changedProperties: PropertyValueMap<unknown> | Map<PropertyKey, unknown>): void {
    super.firstUpdated(changedProperties);
    const parent = this.renderRoot;
    parent.addEventListener('dataChange', (e: CustomEvent<DataChangedDetail>): void => this.handleDataChange(e));
  }

  private handleDataChange(e: CustomEvent<DataChangedDetail>): void {
    const error = e.detail.error ? e.detail.error : null;

    if (error) {
      this.error = error;
      return;
    }

    // if locale is specified, then convert it to lower case
    if (this.locale) {
      this.locale = this.locale.toLowerCase();
    }

    const response = e.detail.response.value;

    // if response is not null and has values, if locale is specified, then
    // get the label in response that has languageTag equal to locale and make it the first label and append the rest of the labels

    const terms = response.map((item: MicrosoftGraph.TermStore.Term) => {
      const labels = item.labels;
      if (labels && labels.length > 0) {
        if (this.locale) {
          const label = labels.find(l => l.languageTag.toLowerCase() === this.locale);
          if (label) {
            item.labels = [label, ...labels.filter(l => l.languageTag.toLowerCase() !== this.locale)];
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

  private handleClick(e: MouseEvent, item: MicrosoftGraph.TermStore.Term) {
    this.selectedTerm = item;
    this.fireCustomEvent('selectionChanged', item);
  }
}
