/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CSSResult, html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { customElement, MgtBaseComponent } from '@microsoft/mgt-element';
import { fluentSearch } from '@fluentui/web-components';
import { registerFluentComponents } from '../../../utils/FluentComponents';
import { strings } from './strings';
import { styles } from './mgt-search-box-css';
import { debounce } from '../../../utils/Utils';

registerFluentComponents(fluentSearch);

/**
 * **Preview component** Web component used to enter a search value to power search scenarios.
 * Component may change before general availability release.
 *
 * @fires {CustomEvent<string>} searchTermChanged - Fired when the search term is changed
 *
 * @class MgtSearchBox
 * @extends {MgtBaseComponent}
 */
@customElement('search-box')
export class MgtSearchBox extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles(): CSSResult[] {
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
   * The placeholder rendered in the search input (for example, `Select a user` or `Select a task list`).
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
   * Debounce delay of the search input in milliseconds
   *
   * @type {number}
   * @memberof MgtSearchBox
   */
  @property({
    attribute: 'debounce-delay',
    type: Number,
    reflect: true
  })
  public debounceDelay: number;
  private _searchTerm = '';
  private debouncedSearchTermChanged: () => void;

  constructor() {
    super();
    console.warn(
      '🦒: <mgt-search-box> is a preview component and may change prior to becoming generally available. See more information https://aka.ms/mgt/preview-components'
    );
    this.debounceDelay = 300;
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
        autocomplete="off"
        class="search-term-input"
        appearance="outline"
        value=${this.searchTerm ?? ''}
        placeholder=${this.placeholder ? this.placeholder : strings.placeholder}
        title=${this.title ? this.title : strings.title}
        @input=${this.onInputChanged}
        @change=${this.onInputChanged}
      >
      </fluent-search>`;
  }

  private readonly onInputChanged = (e: Event) => {
    this.searchTerm = (e.target as HTMLInputElement).value;
  };

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
