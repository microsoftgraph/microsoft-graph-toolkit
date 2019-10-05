/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
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
   * @param template the template to render
   * @param context the data context to be applied
   * @param converters the converter functions used to transform the data
   */
  public static renderTemplate(template: HTMLTemplateElement, context: object, converters?: object) {
    if (template.content && template.content.childNodes.length) {
      const templateContent = template.content.cloneNode(true);
      return this.renderNode(templateContent, context, converters);
    } else if (template.childNodes.length) {
      const div = document.createElement('div');
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < template.childNodes.length; i++) {
        div.appendChild(template.childNodes[i].cloneNode(true));
      }
      return this.renderNode(div, context, converters);
    }
  }
  private static _expression = /{{\s*[\w\.]+\s*}}/g;
  private static _converterExpression = /{{{\s*[\w\.()]+\s*}}}/g;

  /**
   * Gets the value of an expanded key in an object
   *
   * Ex:
   * ```
   * let value = getValueFromObject({d: 3, a: {b: {c: 5}}}, 'a.b.c')
   * ```
   * @param obj the object holding the value (ex: {d: 3, a: {b: {c: 5}}})
   * @param key the key of the value we need (ex: 'a.b.c')
   */
  private static getValueFromObject(obj: object, key: string) {
    const keys = key.trim().split('.');
    let value = obj;
    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < keys.length; i++) {
      const currentKey = keys[i];
      value = value[currentKey];
      if (!value) {
        return null;
      }
    }

    return value;
  }

  private static replaceExpression(str: string, context: object, converters: object) {
    return str
      .replace(this._converterExpression, match => {
        if (!converters) {
          return '';
        }
        return this.evalInContext(match.substring(3, match.length - 3).trim(), { ...converters, ...context });
      })
      .replace(this._expression, match => {
        const key = match.substring(2, match.length - 2);
        const value = this.getValueFromObject(context, key);
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

  private static renderNode(node: Node, context: object, converters: object) {
    if (node.nodeName === '#text') {
      node.textContent = this.replaceExpression(node.textContent, context, converters);
      return node;
    }

    // tslint:disable-next-line: prefer-const
    let nodeElement = node as HTMLElement;

    // replace attribute values
    if (nodeElement.attributes) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < nodeElement.attributes.length; i++) {
        const attribute = nodeElement.attributes[i];
        nodeElement.setAttribute(attribute.name, this.replaceExpression(attribute.value, context, converters));
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
          if (!this.evalBoolInContext(expression, context)) {
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
          this.renderNode(childNode, context, converters);
        }
      } else {
        this.renderNode(childNode, context, converters);
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
      const loopTokens = loopExpression.split(' ');

      if (loopTokens.length > 1) {
        // don't really care what's in the middle at this point
        const itemName = loopTokens[0];
        const listKey = loopTokens[loopTokens.length - 1];

        const list = this.getValueFromObject(context, listKey);
        if (Array.isArray(list)) {
          // first remove the child
          // we will need to make copy of the child for
          // each element in the list
          nodeElement.removeChild(childElement);
          childElement.removeAttribute('data-for');

          for (let j = 0; j < list.length; j++) {
            const newContext: any = {};
            newContext[itemName] = list[j];
            // tslint:disable-next-line: no-string-literal
            newContext.index = j;

            const clone = childElement.cloneNode(true);
            this.renderNode(clone, newContext, converters);
            nodeElement.appendChild(clone);
          }
        }
      }
    }

    return node;
  }

  private static evalBoolInContext(expression, context) {
    return new Function('with(this) { return !!(' + expression + ')}').call(context);
  }

  private static evalInContext(expression, context) {
    const func = new Function('with(this) { return ' + expression + ';}');
    let result;
    try {
      result = func.call(context);
      // tslint:disable-next-line: no-empty
    } catch (e) {}
    return result;
  }
}
