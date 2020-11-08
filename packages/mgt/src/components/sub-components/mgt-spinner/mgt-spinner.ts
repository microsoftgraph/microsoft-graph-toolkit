/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html } from 'lit-element';
import { MgtBaseComponent } from '@microsoft/mgt-element';
import { styles } from './mgt-spinner-css';

/**
 * Custom Component used to handle loading state in components.
 *
 * @export MgtSpinner
 * @class MgtSpinner
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-spinner')
export class MgtSpinner extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  public static get styles() {
    return styles;
  }

  /**
   * Render the loading spinner
   *
   * @returns
   * @memberof MgtSpinner
   */
  public render() {
    return html`
      <div class="spinner"></div>
    `;
  }
}
