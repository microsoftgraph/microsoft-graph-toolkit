/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Helper class for Template Instantiation
 *
 * @export
 * @class TemplateHelper
 */
export class TemplateHelper {
  /**
   * Render a template into a HTMLElement with the appropriate data context
   *
   * Ex:
   * ```
   * <template>
   *  <div>{{myObj.someStr}}</div>
   *  <div data-for="key in myObj.list">
   *    <div>{{key.anotherStr}}</div>
   *  </div>
   * </template>
   * ```
   *
   * @param root the root element to parent the rendered content
   * @param template the template to render
   * @param context the data context to be applied
   */
  public static renderTemplate(root: HTMLElement, template: HTMLTemplateElement, context: object) {
    // inherit context from parent template
    if ((template as any).$parentTemplateContext) {
      context = { ...context, $parent: (template as any).$parentTemplateContext };
    }

    let rendered: Node;

    if (template.content && template.content.childNodes.length) {
      const templateContent = template.content.cloneNode(true);
      rendered = this.renderNode(templateContent, root, context);
    } else if (template.childNodes.length) {
      const div = document.createElement('div');
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < template.childNodes.length; i++) {
        div.appendChild(this.simpleCloneNode(template.childNodes[i]));
      }
      rendered = this.renderNode(div, root, context);
    }

    if (rendered) {
      root.appendChild(rendered);
    }
  }

  /**
   * Set an alternative binding syntax. Default is {{ <value> }}
   *
   * @static
   * @param {string} startStr start of binding syntax
   * @param {string} endStr end of binding syntax
   * @memberof TemplateHelper
   */
  public static setBindingSyntax(startStr: string, endStr: string) {
    this._startExpression = startStr;
    this._endExpression = endStr;

    const start = this.escapeRegex(this._startExpression);
    const end = this.escapeRegex(this._endExpression);

    this._expression = new RegExp(`${start}\\s*([$\\w\\.,'"\\s()\\[\\]]+)\\s*${end}`, 'g');
  }

  /**
   * Global context containing data or functions available to
   * all templates for binding
   *
   * @readonly
   * @static
   * @memberof TemplateHelper
   */
  public static get globalContext() {
    return this._globalContext;
  }

  private static _globalContext = {};

  private static get expression() {
    if (!this._expression) {
      this.setBindingSyntax('{{', '}}');
    }

    return this._expression;
  }

  private static _startExpression: string;
  private static _endExpression: string;
  private static _expression: RegExp;

  private static escapeRegex(string) {
    return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  }

  // simple implementation of deep cloneNode
  // required for nested templates in polyfilled browsers
  private static simpleCloneNode(node: ChildNode) {
    if (!node) {
      return null;
    }

    const clone = node.cloneNode(false);

    // tslint:disable-next-line:prefer-for-of
    for (let i = 0; i < node.childNodes.length; i++) {
      const childClone = this.simpleCloneNode(node.childNodes[i]);
      if (childClone) {
        clone.appendChild(childClone);
      }
    }

    return clone;
  }

  private static expandExpressionsAsString(str: string, context: object) {
    return str.replace(this.expression, (match, p1) => {
      const value = this.evalInContext(p1 || this.trimExpression(match), context);
      if (value) {
        if (typeof value === 'object') {
          return JSON.stringify(value);
        } else {
          return (value as any).toString();
        }
      }
      return '';
    });
  }

  private static renderNode(node: Node, root: HTMLElement, context: object) {
    if (node.nodeName === '#text') {
      node.textContent = this.expandExpressionsAsString(node.textContent, context);
      return node;
    } else if (node.nodeName === 'TEMPLATE') {
      (node as any).$parentTemplateContext = context;
      return node;
    }

    // tslint:disable-next-line: prefer-const
    let nodeElement = node as HTMLElement;

    // replace attribute values
    if (nodeElement.attributes) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < nodeElement.attributes.length; i++) {
        const attribute = nodeElement.attributes[i];

        if (attribute.name === 'data-props') {
          const propsValue = this.trimExpression(attribute.value);
          for (const prop of propsValue.split(',')) {
            const keyValue = prop.trim().split(':');
            if (keyValue.length === 2) {
              const key = keyValue[0].trim();
              const value = this.evalInContext(keyValue[1].trim(), context);

              if (key.startsWith('@')) {
                // event
                if (typeof value === 'function') {
                  nodeElement.addEventListener(key.substring(1), e => value(e, context, root));
                }
              } else {
                nodeElement[key] = value;
              }
            }
          }
        } else {
          nodeElement.setAttribute(attribute.name, this.expandExpressionsAsString(attribute.value, context));
        }
      }
    }

    // don't process nodes that will loop yet, but
    // keep a reference of them
    const loopChildren = [];

    // list of children to remove (ex, when data-if == false)
    const removeChildren = [];
    let previousChildWasIfAndTrue = false;

    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < node.childNodes.length; i++) {
      const childNode = node.childNodes[i];
      const childElement = childNode as HTMLElement;
      let previousChildWasIfAndTrueSet = false;

      if (childElement.dataset) {
        let childWillBeRemoved = false;

        if (childElement.dataset.if) {
          const expression = childElement.dataset.if;
          if (!this.evalBoolInContext(this.trimExpression(expression), context)) {
            removeChildren.push(childElement);
            childWillBeRemoved = true;
          } else {
            childElement.removeAttribute('data-if');
            previousChildWasIfAndTrue = true;
            previousChildWasIfAndTrueSet = true;
          }
        } else if (typeof childElement.dataset.else !== 'undefined') {
          if (previousChildWasIfAndTrue) {
            removeChildren.push(childElement);
            childWillBeRemoved = true;
          } else {
            childElement.removeAttribute('data-else');
          }
        }

        if (childElement.dataset.for && !childWillBeRemoved) {
          loopChildren.push(childElement);
        } else if (!childWillBeRemoved) {
          this.renderNode(childNode, root, context);
        }
      } else {
        this.renderNode(childNode, root, context);
      }

      // clear the flag if the current node wasn't data-if
      // or if it was data-if but it wasn't true
      if (!previousChildWasIfAndTrueSet && childNode.nodeName !== '#text') {
        previousChildWasIfAndTrue = false;
      }
    }

    // now handle nodes that need to be removed
    for (const child of removeChildren) {
      nodeElement.removeChild(child);
    }

    // now handle nodes that should loop
    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < loopChildren.length; i++) {
      const childElement = loopChildren[i] as HTMLElement;

      const loopExpression = childElement.dataset.for;
      const loopTokens = this.trimExpression(loopExpression).split(/\s(in|of)\s/i);

      if (loopTokens.length === 3) {
        // don't really care what's in the middle at this point
        const itemName = loopTokens[0].trim();
        const listKey = loopTokens[2].trim();

        const list = this.evalInContext(listKey, context);
        if (Array.isArray(list)) {
          // first remove the child
          // we will need to make copy of the child for
          // each element in the list
          childElement.removeAttribute('data-for');

          for (let j = 0; j < list.length; j++) {
            const newContext = {
              $index: j,
              ...context
            };
            newContext[itemName] = list[j];

            const clone = childElement.cloneNode(true);
            this.renderNode(clone, root, newContext);
            nodeElement.insertBefore(clone, childElement);
          }
        }
        nodeElement.removeChild(childElement);
      }
    }

    return node;
  }

  private static evalBoolInContext(expression, context) {
    context = { ...context, ...this.globalContext };
    return new Function('with(this) { return !!(' + expression + ')}').call(context);
  }

  private static evalInContext(expression, context) {
    context = { ...context, ...this.globalContext };
    const func = new Function('with(this) { return ' + expression + ';}');
    let result;
    try {
      result = func.call(context);
      // tslint:disable-next-line: no-empty
    } catch (e) {}
    return result;
  }

  private static trimExpression(expression: string) {
    expression = expression.trim();

    if (expression.startsWith(this._startExpression) && expression.endsWith(this._endExpression)) {
      expression = expression.substr(
        this._startExpression.length,
        expression.length - this._startExpression.length - this._endExpression.length
      );
      expression = expression.trim();
    }

    return expression;
  }
}
