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
import { styles } from './mgt-person-card-dummy-css';

/**
 * foo
 *
 * @export
 * @class MgtPersonCardProfile
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-person-card-dummy')
export class MgtPersonCardDummy extends BasePersonCardSection {
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
   * @memberof MgtPersonCardDummy
   */
  public get displayName(): string {
    return 'Dummy';
  }

  /**
   * user-id property allows developer to use id value for component
   * @type {string}
   */
  @property({
    attribute: 'user-id'
  })
  public userId: string;

  constructor() {
    super();

    this.userId = null;
  }

  /**
   * foo
   *
   * @param {string} name
   * @param {string} oldValue
   * @param {string} newValue
   * @returns
   * @memberof MgtPersonCardDummy
   */
  public attributeChangedCallback(name: string, oldValue: string, newValue: string) {
    super.attributeChangedCallback(name, oldValue, newValue);

    if (oldValue === newValue) {
      return;
    }

    switch (name) {
      case 'user-id':
        // reset props here
        this.requestStateUpdate();
        break;
    }
  }

  /**
   * foo
   *
   * @returns {TemplateResult}
   * @memberof MgtPersonCardDummy
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
   * @memberof MgtPersonCardDummy
   */
  public renderIcon(): TemplateResult {
    return html`
      icon
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

    if (!this.userId) {
      return;
    }

    const graph = provider.graph.forComponent(this);
    const betaGraph = BetaGraph.fromGraph(graph);

    const userId = this.userId;
    // const profile = await getProfile(betaGraph, userId);
  }
}
