/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { MgtBaseComponent } from './baseComponent';
import { TemplateHelper } from './templateHelper';

/**
 * An abstract class that defines a templatable web component
 *
 * @export
 * @abstract
 * @class MgtTemplatedComponent
 * @extends {MgtBaseComponent}
 */
export abstract class MgtTemplatedComponent extends MgtBaseComponent {
  /**
   * Collection of functions to be used in template binding
   *
   * @type {*}
   * @memberof MgtTemplatedComponent
   */
  public templateConverters: any = {};

  /**
   * Holds all templates defined by developer
   *
   * @protected
   * @memberof MgtTemplatedComponent
   */
  protected templates = {};

  private _renderedSlots = false;
  private _slotNamesAddedDuringRender = [];

  constructor() {
    super();

    this.templateConverters.lower = (str: string) => str.toLowerCase();
    this.templateConverters.upper = (str: string) => str.toUpperCase();
  }

  /**
   * Updates the element. This method reflects property values to attributes.
   * It can be overridden to render and keep updated element DOM.
   * Setting properties inside this method will *not* trigger
   * another update.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected update(changedProperties) {
    this.templates = this.getTemplates();
    this._slotNamesAddedDuringRender = [];
    super.update(changedProperties);
  }

  /**
   * Invoked whenever the element is updated. Implement to perform
   * post-updating tasks via DOM APIs, for example, focusing an element.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected updated() {
    this.removeUnusedSlottedElements();
  }

  /**
   * Render a <template> by type and return content to render
   *
   * @param templateType type of template (indicated by the data-type attribute)
   * @param context the data context that should be expanded in template
   * @param slotName the slot name that will be used to host the new rendered template. set to a unique value if multiple templates of this type will be rendered. default is templateType
   */
  protected renderTemplate(templateType: string, context: object, slotName?: string) {
    if (!this.templates[templateType]) {
      return null;
    }

    slotName = slotName || templateType;
    this._slotNamesAddedDuringRender.push(slotName);
    this._renderedSlots = true;

    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < this.children.length; i++) {
      if (this.children[i].slot === slotName) {
        return html`
          <slot name=${slotName}></slot>
        `;
      }
    }

    const templateContent = TemplateHelper.renderTemplate(
      this.templates[templateType],
      context,
      this.templateConverters
    );

    const div = document.createElement('div');
    div.slot = slotName;
    div.dataset.generated = 'template';

    if (templateContent) {
      div.appendChild(templateContent);
    }

    this.appendChild(div);

    this.fireCustomEvent('templateRendered', { templateType, context, element: div });

    return html`
      <slot name=${slotName}></slot>
    `;
  }

  private getTemplates() {
    const templates: any = {};

    // tslint:disable-next-line: prefer-for-of
    for (let i = 0; i < this.children.length; i++) {
      const child = this.children[i];
      if (child.nodeName === 'TEMPLATE') {
        const template = child as HTMLElement;
        if (template.dataset.type) {
          templates[template.dataset.type] = template;
        } else {
          templates.default = template;
        }
      }
    }

    return templates;
  }

  private removeUnusedSlottedElements() {
    if (this._renderedSlots) {
      for (let i = 0; i < this.children.length; i++) {
        const child = this.children[i] as HTMLElement;
        if (child.dataset && child.dataset.generated && !this._slotNamesAddedDuringRender.includes(child.slot)) {
          this.removeChild(child);
          i--;
        }
      }
      this._renderedSlots = false;
    }
  }
}
