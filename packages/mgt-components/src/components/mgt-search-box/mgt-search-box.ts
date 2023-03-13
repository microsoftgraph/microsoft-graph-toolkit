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

registerFluentComponents(fluentSearch);

/**
 * Web component used to enter a search value to power search scenarios
 *
 * @fires {CustomEvent<string>} searchTermChanged - Fired when the search term is changed
 *
 * @class MgtSearchBox
 * @extends {MgtBaseComponent}
 */
@customElement('search-box')
class MgtSearchBox extends MgtBaseComponent {
  constructor() {
    super();
  }

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
   * Value of the search input
   *
   * @type {string}
   * @memberof MgtSearchBox
   */
  @property({
    attribute: 'value',
    type: String
  })
  public value: string;

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
        autofocus
        placeholder=${this.placeholder ? this.placeholder : strings.placeholder}
        @input=${this.onInputChanged}
      >
      </fluent-search>
`;
  }

  /**
   * Fires the searchTermChanged even when value changes
   * @param e
   */
  private onInputChanged(e: Event) {
    this.value = (e.target as HTMLInputElement).value;
    this.fireCustomEvent('searchTermChanged', this.value);
  }
}
