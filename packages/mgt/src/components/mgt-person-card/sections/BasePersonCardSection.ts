/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property, TemplateResult } from 'lit-element';
import { MgtTemplatedComponent } from '../../../components/templatedComponent';
import { IDynamicPerson } from '../../../graph/types';
import { MgtPersonCard } from '../mgt-person-card';

/**
 * foo
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
   * foo
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
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  public abstract renderIcon(): TemplateResult;

  /**
   * foo
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
   * foo
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
   * foo
   *
   * @protected
   * @abstract
   * @memberof BasePersonCardSection
   */
  public abstract clearState(): void;

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    return this.isCompact ? this.renderCompactView() : this.renderFullView();
  }

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  protected abstract renderCompactView(): TemplateResult;

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  protected abstract renderFullView(): TemplateResult;

  /**
   * foo
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
