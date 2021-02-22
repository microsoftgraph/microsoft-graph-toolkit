/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtTemplatedComponent } from '@microsoft/mgt-element';
import { html, property, TemplateResult } from 'lit-element';

import { IDynamicPerson } from '../../../graph/types';
import { MgtPersonCard } from '../mgt-person-card';

import '../../sub-components/mgt-spinner/mgt-spinner';

/**
 * A base class for building person card subsections.
 *
 * @export
 * @class BasePersonCardSection
 * @extends {MgtTemplatedComponent}
 */
export abstract class BasePersonCardSection extends MgtTemplatedComponent {
  /**
   * Set the person details to render
   *
   * @type {IDynamicPerson}
   * @memberof BasePersonCardSection
   */
  @property({
    attribute: 'person-details',
    type: Object
  })
  public get personDetails(): IDynamicPerson {
    return this._personDetails;
  }
  public set personDetails(value: IDynamicPerson) {
    if (this._personDetails === value) {
      return;
    }

    this._personDetails = value;
    this.requestStateUpdate();
  }

  /**
   * The name for display in the overview section.
   *
   * @abstract
   * @type {string}
   * @memberof BasePersonCardSection
   */
  public abstract get displayName(): string;

  /**
   * Determines the appropriate view state: full or compact
   *
   * @protected
   * @type {boolean}
   * @memberof BasePersonCardSection
   */
  protected get isCompact(): boolean {
    return this._isCompact;
  }

  private _isCompact: boolean;
  private _personDetails: IDynamicPerson;

  constructor() {
    super();
    this._isCompact = false;
    this._personDetails = null;
  }

  /**
   * Render the icon for display in the navigation ribbon.
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  public abstract renderIcon(): TemplateResult;

  /**
   * Set the section to compact view mode
   *
   * @returns
   * @memberof BasePersonCardSection
   */
  public asCompactView() {
    this._isCompact = true;
    this.requestUpdate();
    return this;
  }

  /**
   * Set the section to full view mode
   *
   * @returns
   * @memberof BasePersonCardSection
   */
  public asFullView() {
    this._isCompact = false;
    this.requestUpdate();
    return this;
  }

  /**
   * Reset any state in the section
   *
   * @protected
   * @abstract
   * @memberof BasePersonCardSection
   */
  public clearState(): void {
    this._isCompact = false;
    this._personDetails = null;
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return this.isCompact ? this.renderCompactView() : this.renderFullView();
  }

  /**
   * Render a spinner while the component loads state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  protected renderLoading(): TemplateResult {
    return html`
      <div class="loading">
        <mgt-spinner></mgt-spinner>
      </div>
    `;
  }

  /**
   * Render the section in a empty data state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPersonCardContact
   */
  protected renderNoData(): TemplateResult {
    return html`
      <div class="no-data">No data</div>
    `;
  }

  /**
   * Render the compact view
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  protected abstract renderCompactView(): TemplateResult;

  /**
   * Render the full view
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  protected abstract renderFullView(): TemplateResult;

  /**
   * Navigate the card to a different user.
   *
   * @protected
   * @memberof BasePersonCardSection
   */
  protected navigateCard(person: IDynamicPerson): void {
    // Search for card parent and update it's personDetails object
    let parent: any = this.parentNode;
    while (parent) {
      parent = parent.parentNode;

      if (parent && parent.host && parent.host.tagName === 'MGT-PERSON-CARD') {
        parent = parent.host;
        break;
      }
    }

    const personCard = parent as MgtPersonCard;
    personCard.navigate(person);
  }
}
