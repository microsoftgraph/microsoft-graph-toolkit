/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property, state } from 'lit/decorators.js';
import { html, PropertyValueMap, PropertyValues, TemplateResult } from 'lit';

import { equals } from '../utils/equals';
import { MgtBaseComponent } from './baseComponent';
import { TemplateContext } from '../utils/TemplateContext';
import { TemplateHelper } from '../utils/TemplateHelper';

/**
 * Lookup for rendered component templates and contexts by slot name.
 */
type RenderedTemplates = Record<
  string,
  {
    /**
     * Reference to the data context used to render the slot.
     */
    context: Record<string, unknown>;
    /**
     * Reference to the rendered DOM element corresponding to the slot.
     */
    slot: HTMLElement;
  }
>;

export interface TemplateRenderedData {
  templateType: string;
  context: Record<string, unknown>;
  element: HTMLElement;
}

type OrderedHtmlTemplate = HTMLTemplateElement & { templateOrder: number };

/**
 * An abstract class that defines a templatable web component
 *
 * @export
 * @abstract
 * @class MgtTemplatedComponent
 * @extends {MgtBaseComponent}
 *
 * @fires {CustomEvent<MgtElement.TemplateRenderedData>} templateRendered - fires when a template is rendered
 */
export abstract class MgtTemplatedComponent extends MgtBaseComponent {
  /**
   * Additional data context to be used in template binding
   * Use this to add event listeners or value converters
   *
   * @type {MgtElement.TemplateContext}
   * @memberof MgtTemplatedComponent
   */
  @property({ attribute: false }) public templateContext: TemplateContext;

  /**
   *
   * Gets or sets the error (if any) of the request
   *
   * @type object
   * @memberof MgtSearchResults
   */
  @state() protected error: object;

  /**
   * Holds all templates defined by developer
   *
   * @protected
   * @memberof MgtTemplatedComponent
   */
  protected templates: Record<string, OrderedHtmlTemplate> = {};

  private _renderedSlots = false;
  private _renderedTemplates: RenderedTemplates = {};
  private _slotNamesAddedDuringRender = [];

  constructor() {
    super();

    this.templateContext = this.templateContext || {};
  }

  /**
   * Updates the element. This method reflects property values to attributes.
   * It can be overridden to render and keep updated element DOM.
   * Setting properties inside this method will *not* trigger
   * another update.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected update(changedProperties: PropertyValueMap<unknown> | Map<PropertyKey, unknown>) {
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
   * * @param changedProperties Map of changed properties with old values
   */
  protected updated(changedProperties: PropertyValues) {
    super.updated(changedProperties);
    this.removeUnusedSlottedElements();
  }

  /**
   * Render a <template> by type and return content to render
   *
   * @param templateType type of template (indicated by the data-type attribute)
   * @param context the data context that should be expanded in template
   * @param slotName the slot name that will be used to host the new rendered template. set to a unique value if multiple templates of this type will be rendered. default is templateType
   */
  protected renderTemplate(templateType: string, context: object, slotName?: string): TemplateResult {
    if (!this.hasTemplate(templateType)) {
      return null;
    }

    slotName = slotName || templateType;
    this._slotNamesAddedDuringRender.push(slotName);
    this._renderedSlots = true;

    const template = html`
      <slot name=${slotName}></slot>
    `;

    const dataContext = { ...context, ...this.templateContext };

    if (Object.prototype.hasOwnProperty.call(this._renderedTemplates, slotName)) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
      const { context: existingContext, slot } = this._renderedTemplates[slotName];
      if (equals(existingContext, dataContext)) {
        return template;
      }
      this.removeChild(slot);
    }

    const div = document.createElement('div');
    div.slot = slotName;
    div.dataset.generated = 'template';

    TemplateHelper.renderTemplate(div, this.templates[templateType], dataContext);

    this.appendChild(div);

    this._renderedTemplates[slotName] = { context: dataContext, slot: div };

    const templateRenderedData: TemplateRenderedData = { templateType, context: dataContext, element: div };
    this.fireCustomEvent('templateRendered', templateRenderedData);

    return template;
  }

  /**
   * Check if a specific template has been provided.
   *
   * @protected
   * @param {string} templateName
   * @returns {boolean}
   * @memberof MgtTemplatedComponent
   */
  protected hasTemplate(templateName: string): boolean {
    return Boolean(this.templates?.[templateName]);
  }

  private getTemplates() {
    const templates: Record<string, OrderedHtmlTemplate> = {};

    for (let i = 0; i < this.children.length; i++) {
      const child = this.children[i];
      if (child.nodeName === 'TEMPLATE') {
        const template = child as OrderedHtmlTemplate;
        if (template.dataset.type) {
          templates[template.dataset.type] = template;
        } else {
          templates.default = template;
        }

        template.templateOrder = i;
      }
    }

    return templates;
  }

  /**
   * Renders an error
   *
   * @returns
   */
  protected renderError(): TemplateResult {
    if (this.hasTemplate('error')) {
      return this.renderTemplate('error', this.error);
    }

    return html`
      <div class="error">
        ${this.error}
      </div>
    `;
  }

  private removeUnusedSlottedElements() {
    if (this._renderedSlots) {
      for (let i = 0; i < this.children.length; i++) {
        const child = this.children[i] as HTMLElement;
        if (child.dataset?.generated && !this._slotNamesAddedDuringRender.includes(child.slot)) {
          this.removeChild(child);
          delete this._renderedTemplates[child.slot];
          i--;
        }
      }
      this._renderedSlots = false;
    }
  }
}
