/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { customElement, MgtBaseComponent } from '@microsoft/mgt-element';
import { fluentSearch } from '@fluentui/web-components/dist/esm/search';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { strings } from './strings';
import { styles } from './mgt-search-box-css';
import { debounce } from '../../utils/Utils';

registerFluentComponents(fluentSearch);

/**
 * Web component used to enter a search value to power search scenarios
 *
 * @fires {CustomEvent<string>} searchTermChanged - Fired when the search term is changed
 *
 * @cssprop --search-input-width - {Length} Search input width
 *
 * @class MgtSearchBox
 * @extends {MgtBaseComponent}
 */
@customElement('search-box')
class MgtSearchBox extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * Provides strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtSearchBox
   */
  protected get strings() {
    return strings;
  }

  /**
   * Placeholder text
   *
   * @type {string}
   * @memberof MgtSearchBox
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Value of the search term
   *
   * @type {string}
   * @memberof MgtSearchBox
   */
  @property({
    attribute: 'search-term',
    type: String
  })
  public get searchTerm() {
    return this._searchTerm;
  }
  public set searchTerm(value) {
    this._searchTerm = value;
    this.fireSearchTermChanged();
  }

  /**
   * Debounce delay of the search input
   *
   * @type {number}
   * @memberof MgtSearchBox
   */
  @property({
    attribute: 'debounce-delay',
    type: Number,
    reflect: true
  })
  public get debounceDelay() {
    return this._debounceDelay;
  }
  public set debounceDelay(value) {
    this._debounceDelay = value;
  }

  private _debounceDelay: number = 300;
  private _searchTerm: string = '';
  private debouncedSearchTermChanged;

  constructor() {
    super();
  }
  /**
   * Renders the component
   *
   * @return {TemplateResult}
   * @memberof MgtSearchBox
   */
  render(): TemplateResult {
    return html`
      <fluent-search
        class="search-term-input"
        appearance="outline"
        value=${this.searchTerm ?? this.searchTerm}
        placeholder=${this.placeholder ? this.placeholder : strings.placeholder}
        title="${strings.placeholder}"
        @input=${(e: Event) => this.onInputChanged(e)}
        @change=${(e: Event) => this.onInputChanged(e)}
      >
      </fluent-search>`;
  }

  private onInputChanged(e: Event) {
    this.searchTerm = (e.target as HTMLInputElement).value;
  }

  /**
   * Fires and debounces the custom event to listeners
   */
  private fireSearchTermChanged() {
    if (!this.debouncedSearchTermChanged) {
      this.debouncedSearchTermChanged = debounce(() => {
        this.fireCustomEvent('searchTermChanged', this.searchTerm);
      }, this.debounceDelay);
    }

    this.debouncedSearchTermChanged();
  }
}
