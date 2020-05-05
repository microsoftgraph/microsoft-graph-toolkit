import { TemplateResult } from 'lit-element';
import { MgtTemplatedComponent } from '../../../../components/templatedComponent';

/**
 * foo
 *
 * @export
 * @class BasePersonCardSection
 * @extends {MgtTemplatedComponent}
 */
export abstract class BasePersonCardSection extends MgtTemplatedComponent {
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
