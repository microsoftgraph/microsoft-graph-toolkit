/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, TemplateResult } from 'lit-element';
import { Providers } from '../../../../../Providers';
import { ProviderState } from '../../../../../providers/IProvider';
import { BetaGraph } from '../../../../BetaGraph';
import { BasePersonCardSection } from '../BasePersonCardSection';
import { styles } from './mgt-person-card-organization-css';

/**
 * foo
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-organization')
export class MgtPersonCardOrganization extends BasePersonCardSection {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * foo
   *
   * @readonly
   * @type {string}
   * @memberof MgtPersonCardOrganization
   */
  public get displayName(): string {
    return 'Organization';
  }

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  public renderCompactView(): TemplateResult {
    return html`
      compact
    `;
  }

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardOrganization
   */
  public renderIcon(): TemplateResult {
    return html`
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M13 4H8V7H13V4ZM8 3C7.44772 3 7 3.44772 7 4V7C7 7.55228 7.44772 8 8 8H10V9H7.5C6.67157 9 6 9.67157 6 10.5V11H4C3.44772 11 3 11.4477 3 12V15C3 15.5523 3.44772 16 4 16H9C9.55228 16 10 15.5523 10 15V12C10 11.4477 9.55228 11 9 11H7V10.5C7 10.2239 7.22386 10 7.5 10H13.5C13.7761 10 14 10.2239 14 10.5V11H12C11.4477 11 11 11.4477 11 12V15C11 15.5523 11.4477 16 12 16H17C17.5523 16 18 15.5523 18 15V12C18 11.4477 17.5523 11 17 11H15V10.5C15 9.67157 14.3284 9 13.5 9H11V8H13C13.5523 8 14 7.55228 14 7V4C14 3.44772 13.5523 3 13 3H8ZM9 12H4L4 15H9V12ZM12 12H17V15H12V12Z"
          fill="#605E5C"
        />
      </svg>
    `;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return html``;
  }

  /**
   * load state into the component
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPersonCardProfile
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;

    // check if user is signed in
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (!this.personDetails) {
      return;
    }

    const graph = provider.graph.forComponent(this);
    const betaGraph = BetaGraph.fromGraph(graph);

    // const userId = this.personDetails.id;
    // const profile = await getProfile(betaGraph, userId);

    // this.profile = profile;
    this.requestUpdate();
  }
}
