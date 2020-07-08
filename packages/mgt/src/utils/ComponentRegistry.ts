import { customElement } from 'lit-element';

/**
 * Manages the registration of custom elements.
 *
 * @export
 * @class ComponentRegistry
 */
export class ComponentRegistry {
  /**
   * Maintains the collection of custom elements available in the toolkit.
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
   * Set a prefix to alias built-in components.
   *
   * JS -> ...usePrefix('foo');
   * HTML -> <foo-mgt-login></foo-mgt-login>
   *
   * @static
   * @param {string} prefix
   * @memberof ComponentRegistry
   */
  public static usePrefix(prefix: string) {
    if (!prefix) {
      return;
    }

    // tslint:disable-next-line: forin
    for (const tagName in this._scopedElements) {
      // External tag name
      const externalTag = `${prefix}-${tagName}`;
      const component = this._scopedElements[tagName];

      // Check for existing registration
      if (!customElements.get(externalTag)) {
        // Create a type clone
        // tslint:disable-next-line: max-classes-per-file
        const clone: any = class extends component {
          constructor() {
            super();
          }
        };

        // Register component
        customElement(externalTag)(clone);
      }
    }
  }

  /**
   * Register a custom component for use in the DOM.
   *
   * @static
   * @param {string} tagName
   * @param {*} component
   * @memberof ComponentRegistry
   */
  public static register(tagName: string, component: any): void {
    // Add the component to the scopedElements object.
    ComponentRegistry._scopedElements[tagName] = component;

    // Check for existing registration
    if (!customElements.get(tagName)) {
      // Register component
      customElement(tagName)(component);
    }
  }

  private static _scopedElements: object = {};
}

/**
 * Decorator for registering custom components
 *
 * @export
 * @param {string} tagName
 * @returns
 */
export function registeredComponent(tagName: string) {
  return component => {
    ComponentRegistry.register(tagName, component);
  };
}
