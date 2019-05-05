/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

export class TemplateHelper {
  private static _expression = /{{\s*[\w\.]+\s*}}/g;

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
    let keys = key.trim().split('.');
    let value = obj;
    for (let i = 0; i < keys.length; i++) {
      let currentKey = keys[i];
      value = value[currentKey];
      if (!value) {
        return null;
      }
    }

    return value;
  }

  private static replaceExpression(str: string, context: object) {
    return str.replace(this._expression, match => {
      let key = match.substring(2, match.length - 2);
      let value = this.getValueFromObject(context, key);
      if (value) {
        if (typeof value == 'object') {
          return JSON.stringify(value);
        } else {
          return (<any>value).toString();
        }
      }
      return '';
    });
  }

  private static renderNode(node: Node, context: object) {
    if (node.nodeName === '#text') {
      node.textContent = this.replaceExpression(node.textContent, context);
      return node;
    }

    let nodeElement = <HTMLElement>node;

    // replace attribute values
    if (nodeElement.attributes) {
      for (let i = 0; i < nodeElement.attributes.length; i++) {
        let attribute = nodeElement.attributes[i];
        nodeElement.setAttribute(attribute.name, this.replaceExpression(attribute.value, context));
      }
    }

    // don't process nodes that will loop yet, but
    // keep a reference of them
    let loopChildren = [];

    // list of children to remove (ex, when data-if == false)
    let removeChildren = [];
    let previousChildWasIfAndTrue = false;

    for (let i = 0; i < node.childNodes.length; i++) {
      let childNode = node.childNodes[i];
      let childElement = <HTMLElement>childNode;
      let previousChildWasIfAndTrueSet = false;

      if (childElement.dataset) {
        let childWillBeRemoved = false;

        if (childElement.dataset.if) {
          let expression = childElement.dataset.if;
          if (!this.evalInContext(expression, context)) {
            removeChildren.push(childElement);
            childWillBeRemoved = true;
          } else {
            childElement.removeAttribute('data-if');
            previousChildWasIfAndTrue = true;
            previousChildWasIfAndTrueSet = true;
          }
        } else if (typeof childElement.dataset.else != 'undefined') {
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
          this.renderNode(childNode, context);
        }
      } else {
        this.renderNode(childNode, context);
      }

      // clear the flag if the current node wasn't data-if
      // or if it was data-if but it wasn't true
      if (!previousChildWasIfAndTrueSet && childNode.nodeName !== '#text') {
        previousChildWasIfAndTrue = false;
      }
    }

    // now handle nodes that need to be removed
    for (let i = 0; i < removeChildren.length; i++) {
      nodeElement.removeChild(removeChildren[i]);
    }

    // now handle nodes that should loop
    for (let i = 0; i < loopChildren.length; i++) {
      let childElement = <HTMLElement>loopChildren[i];

      let loopExpression = childElement.dataset.for;
      let loopTokens = loopExpression.split(' ');

      if (loopTokens.length > 1) {
        // don't really care what's in the middle at this point
        let itemName = loopTokens[0];
        let listKey = loopTokens[loopTokens.length - 1];

        let list = this.getValueFromObject(context, listKey);
        if (Array.isArray(list)) {
          // first remove the child
          // we will need to make copy of the child for
          // each element in the list
          nodeElement.removeChild(childElement);
          childElement.removeAttribute('data-for');

          for (let j = 0; j < list.length; j++) {
            let newContext = {};
            newContext[itemName] = list[j];
            newContext['index'] = j;

            let clone = childElement.cloneNode(true);
            this.renderNode(clone, newContext);
            nodeElement.appendChild(clone);
          }
        }
      }
    }

    return node;
  }

  static evalInContext(expression, context) {
    return new Function('with(this) { return !!(' + expression + ')}').call(context);
  }

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
   */
  public static renderTemplate(template: HTMLTemplateElement, context: object) {
    if (template.content && template.content.childNodes.length) {
      let templateContent = template.content.cloneNode(true);
      return this.renderNode(templateContent, context);
    } else if (template.childNodes.length) {
      let div = document.createElement('div');
      for (let i = 0; i < template.childNodes.length; i++) {
        div.appendChild(template.childNodes[i].cloneNode(true));
      }
      return this.renderNode(div, context);
    }
  }
}
