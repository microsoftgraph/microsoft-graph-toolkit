/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, property, LitElement, html } from 'lit-element';

/**
 * Localizes component strings with definied user agent strings
 *
 *
 * @export
 * @class LocalizationService
 */
@customElement('mgt-localization-service')
export class MgtLocalizationService extends LitElement {
  @property({
    attribute: 'strings',
    type: String
  })
  public strings: string;

  updated() {
    let event = new CustomEvent('strings', {
      detail: { strings: this.strings },
      bubbles: true,
      composed: true
    });
    this.dispatchEvent(event);
  }
}
