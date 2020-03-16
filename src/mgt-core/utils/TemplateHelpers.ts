/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

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
 * @param additionalContext additional context that could contain functions to transform the data
 */
export function renderTemplate(
  root: HTMLElement,
  template: HTMLTemplateElement,
  context: object,
  additionalContext?: object
) {
  // inherit context from parent template
  if ((template as any).$parentTemplateContext) {
    context = { ...context, $parent: (template as any).$parentTemplateContext };
  }

  let rendered: Node;

  if (template.content && template.content.childNodes.length) {
    const templateContent = template.content.cloneNode(true);
    rendered = renderNode(templateContent, root, context, additionalContext);
  } else if (template.childNodes.length) {
    const div = document.createElement('div');
    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < template.childNodes.length; i++) {
      div.appendChild(template.childNodes[i].cloneNode(true));
    }
    rendered = renderNode(div, root, context, additionalContext);
  }

  if (rendered) {
    root.appendChild(rendered);
  }
}

// tslint:disable-next-line: completed-docs
const _expression = /{{+\s*[$\w\.()\[\]]+\s*}}+/g;

// tslint:disable-next-line: completed-docs
function expandExpressionsAsString(str: string, context: object, additionalContext: object) {
  return str.replace(_expression, match => {
    const value = evalInContext(trimExpression(match), { ...context, ...additionalContext });
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

// tslint:disable-next-line: completed-docs
function renderNode(node: Node, root: HTMLElement, context: object, additionalContext: object) {
  if (node.nodeName === '#text') {
    node.textContent = expandExpressionsAsString(node.textContent, context, additionalContext);
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
        const propsValue = trimExpression(attribute.value);
        for (const prop of propsValue.split(',')) {
          const keyValue = prop.trim().split(':');
          if (keyValue.length === 2) {
            const key = keyValue[0].trim();
            const value = evalInContext(keyValue[1].trim(), { ...context, ...additionalContext });

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
        nodeElement.setAttribute(
          attribute.name,
          expandExpressionsAsString(attribute.value, context, additionalContext)
        );
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
        if (!evalBoolInContext(trimExpression(expression), { ...context, ...additionalContext })) {
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
        renderNode(childNode, root, context, additionalContext);
      }
    } else {
      renderNode(childNode, root, context, additionalContext);
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
    const loopTokens = trimExpression(loopExpression).split(/\s+(in|of)\s+/i);

    if (loopTokens.length === 3) {
      // don't really care what's in the middle at this point
      const itemName = loopTokens[0];
      const listKey = loopTokens[2];

      const list = evalInContext(listKey, { ...context, ...additionalContext });
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
          renderNode(clone, root, newContext, additionalContext);
          nodeElement.insertBefore(clone, childElement);
        }
      }
      nodeElement.removeChild(childElement);
    }
  }

  return node;
}

// tslint:disable-next-line: completed-docs
function evalBoolInContext(expression, context) {
  return new Function('with(this) { return !!(' + expression + ')}').call(context);
}

// tslint:disable-next-line: completed-docs
function evalInContext(expression, context) {
  const func = new Function('with(this) { return ' + expression + ';}');
  let result;
  try {
    result = func.call(context);
    // tslint:disable-next-line: no-empty
  } catch (e) {}
  return result;
}

// tslint:disable-next-line: completed-docs
function trimExpression(expression: string) {
  let start = 0;
  let end = expression.length - 1;

  while (expression[start] === '{' && start < end) {
    start++;
  }

  while (expression[end] === '}' && start <= end) {
    end--;
  }

  return expression.substring(start, end + 1).trim();
}
