import { property, TemplateResult } from 'lit-element';
import { MgtTemplatedComponent } from '../../../../components/templatedComponent';
import { IDynamicPerson } from '../../../../graph/types';

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

  private _personDetails: IDynamicPerson;

  /**
   * foo
   *
   * @abstract
   * @type {string}
   * @memberof BasePersonCardSection
   */
  public abstract get displayName(): string;

  /**
   * foo
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  public abstract renderCompactView(): TemplateResult;

  /**
   * foo
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof BasePersonCardSection
   */
  public abstract renderIcon(): TemplateResult;
}
