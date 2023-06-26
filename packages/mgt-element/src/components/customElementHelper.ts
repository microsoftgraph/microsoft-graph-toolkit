/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

class CustomElementHelper {
  private get defaultPrefix() {
    return 'mgt';
  }
  private _disambiguation = '';

  /**
   * Adds a disambiguation segment to the custom elements registered with the browser
   *
   * @param {string} disambiguation
   * @return {CustomElementHelper} the current object
   * @memberof CustomElementHelper
   */
  public withDisambiguation(disambiguation: string) {
    if (disambiguation && !this._disambiguation) this._disambiguation = disambiguation;
    return this;
  }

  /**
   * Provides the prefix to be used for mgt web component tags
   *
   * @readonly
   * @type {string}
   * @memberof CustomElementHelper
   */
  public get prefix(): string {
    return this._disambiguation ? `${this.defaultPrefix}-${this._disambiguation}` : this.defaultPrefix;
  }

  /**
   * Returns the current value for the disambiguation
   *
   * @readonly
   * @type {string}
   * @memberof CustomElementHelper
   */
  public get disambiguation(): string {
    return this._disambiguation;
  }

  /**
   * Returns true if a value has been provided for the disambiguation
   *
   * @readonly
   * @type {boolean}
   * @memberof CustomElementHelper
   */
  public get isDisambiguated(): boolean {
    return Boolean(this._disambiguation);
  }

  /**
   * Removes disambiguation from the provided tagName.
   * Intended for use when providing tag names in analytics headers passed by the Graph client
   *
   * @param {string} tagName
   * @return {*}  {string}
   * @memberof CustomElementHelper
   */
  public normalize(tagName: string): string {
    return this.isDisambiguated
      ? tagName.toUpperCase().replace(this.prefix.toUpperCase(), this.defaultPrefix.toUpperCase())
      : tagName;
  }
}

/**
 * Helper object to provide the desired prefix for mgt web component tags
 *
 * @type CustomElementHelper
 */
const customElementHelper = new CustomElementHelper();
export { customElementHelper };
