import { customElement } from 'lit-element';

/**
 * foo
 *
 * @export
 * @class ComponentRegistry
 */
export class ComponentRegistry {
  /**
   * foo
   *
   * @readonly
   * @static
   * @type {object}
   * @memberof ComponentRegistry
   */
  public static get scopedElements(): object {
    return this._scopedElements;
  }

  /**
   * foo
   *
   * @static
   * @param {string} prefix
   * @memberof ComponentRegistry
   */
  public static setPrefix(prefix: string) {
    this._prefix = prefix;
  }

  /**
   * foo
   *
   * @static
   * @param {string} tagName
   * @param {*} component
   * @memberof ComponentRegistry
   */
  public static register(tagName: string, component: any): void {
    // Add the component to the scopedElements object.
    ComponentRegistry._scopedElements[tagName] = component;

    // External tag name
    const externalTag = this._prefix ? `${this._prefix}-${tagName}` : tagName;

    // Check for existing registration
    if (!customElements.get(externalTag)) {
      // Register component
      customElement(externalTag)(component);
    }
  }

  private static _prefix: string;
  private static _scopedElements: object = {};
}
