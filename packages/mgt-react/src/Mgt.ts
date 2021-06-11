/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React, { ReactNode, ReactElement } from 'react';
import ReactDOM from 'react-dom';
import { Wc, WcProps, WcTypeProps } from 'wc-react';

export class Mgt extends Wc {
  private _templates: Record<string, ReactElement>;

  constructor(props: WcTypeProps) {
    super(props);

    this.handleTemplateRendered = this.handleTemplateRendered.bind(this);
  }

  protected getTag() {
    let tag = super.getTag();
    if (!tag.startsWith('mgt-')) {
      tag = 'mgt-' + tag;
    }

    return tag;
  }

  public render() {
    const tag = this.getTag();
    if (!tag) {
      throw '"wcType" must be set!';
    }

    this.processTemplates(this.props.children);

    let templateElements = [];

    if (this._templates) {
      for (let t in this._templates) {
        if (this._templates.hasOwnProperty(t)) {
          const element = React.createElement('template', { key: t, 'data-type': t }, null);
          templateElements.push(element);
        }
      }
    }

    return React.createElement(tag, { ref: (element: HTMLElement) => this.setRef(element) }, templateElements);
  }

  /**
   * Sets the web component reference and syncs the props
   *
   * @protected
   * @param {HTMLElement} element
   * @memberof Wc
   */
  protected setRef(component: HTMLElement) {
    if (component) {
      component.addEventListener('templateRendered', this.handleTemplateRendered);
    }
    super.setRef(component);
  }

  /**
   * Removes all event listeners from web component element
   *
   * @protected
   * @returns
   * @memberof Mgt
   */
  protected cleanUp() {
    if (!this.element) {
      return;
    }

    this.element.removeEventListener('templateRendered', this.handleTemplateRendered);

    super.cleanUp();
  }

  /**
   * Renders a template
   *
   * @protected
   * @param {*} e
   * @returns
   * @memberof Mgt
   */
  protected handleTemplateRendered(e) {
    if (!this._templates) {
      return;
    }

    let templateType = e.detail.templateType;
    let dataContext = e.detail.context;
    let element = e.detail.element;

    let template = this._templates[templateType];

    if (template) {
      template = React.cloneElement(template, { dataContext });
      ReactDOM.render(template, element);
    }
  }

  /**
   * Prepares templates for rendering
   *
   * @protected
   * @param {ReactNode} children
   * @returns
   * @memberof Mgt
   */
  protected processTemplates(children: ReactNode) {
    if (!children) {
      return;
    }

    let templates = {};

    React.Children.forEach(children, child => {
      let element = child as ReactElement;
      if (element && element.props && element.props.template) {
        templates[element.props.template] = element;
      } else {
        templates['default'] = element;
      }
    });

    this._templates = templates;
  }
}

/**
 * Creates a new React Functional Component that wraps the
 * web component with the specified tag name
 *
 * @template T - optional props type for component
 * @param {(string | Function)} tag
 * @returns React component
 */
export const wrapMgt = <T = WcProps>(tag: string) => {
  const component: React.ForwardRefExoticComponent<
    React.PropsWithoutRef<T & React.HTMLAttributes<any>> & React.RefAttributes<unknown>
  > = React.forwardRef((props: T, ref) => React.createElement(Mgt, { wcType: tag, innerRef: ref, ...props }));
  return component;
};
