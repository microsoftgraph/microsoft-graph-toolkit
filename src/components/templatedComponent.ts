/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit-element';
import { TemplateHelper } from './templateHelper';
import { MgtBaseComponent } from './baseComponent';

export abstract class MgtTemplatedComponent extends MgtBaseComponent {
  private _renderedSlots = false;
  private _slotNamesAddedDuringRender = [];
  protected templates = {};

  public templateConverters: any = {};

  constructor() {
    super();

    this.templateConverters.lower = (str: string) => str.toLowerCase();
    this.templateConverters.upper = (str: string) => str.toUpperCase();
  }

  protected update(changedProperties) {
    this.templates = this.getTemplates();
    this._slotNamesAddedDuringRender = [];
    super.update(changedProperties);
  }

  protected updated() {
    this.removeUnusedSlottedElements();
  }

  private getTemplates() {
    let templates = {};

    for (let i = 0; i < this.children.length; i++) {
      let child = this.children[i];
      if (child.nodeName == 'TEMPLATE') {
        let template = <HTMLElement>child;
        if (template.dataset.type) {
          templates[template.dataset.type] = template;
        } else {
          templates['default'] = template;
        }
      }
    }

    return templates;
  }

  private removeUnusedSlottedElements() {
    if (this._renderedSlots) {
      for (let i = 0; i < this.children.length; i++) {
        let child = <HTMLElement>this.children[i];
        if (child.dataset && child.dataset.generated && !this._slotNamesAddedDuringRender.includes(child.slot)) {
          this.removeChild(child);
          i--;
        }
      }
      this._renderedSlots = false;
    }
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

    for (let i = 0; i < this.children.length; i++) {
      if (this.children[i].slot == slotName) {
        return html`
          <slot name=${slotName}></slot>
        `;
      }
    }

    let templateContent = TemplateHelper.renderTemplate(this.templates[templateType], context, this.templateConverters);

    let div = document.createElement('div');
    div.slot = slotName;
    div.dataset.generated = 'template';
    div.appendChild(templateContent);

    this.appendChild(div);

    this.fireCustomEvent('templateRendered', { templateType: templateType, context: context, element: div });

    return html`
      <slot name=${slotName}></slot>
    `;
  }
}
