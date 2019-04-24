import { LitElement, html } from 'lit-element';
import { TemplateHelper } from './templateHelper';

export abstract class MgtTemplatedComponent extends LitElement {
  private fireCustomEvent(eventName: string, detail?: any): boolean {
    let event = new CustomEvent(eventName, {
      cancelable: true,
      bubbles: false,
      detail: detail
    });
    return this.dispatchEvent(event);
  }

  getTemplates() {
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

  removeSlottedElements() {
    for (let i = 0; i < this.children.length; i++) {
      let child = <HTMLElement>this.children[i];
      if (child.dataset && child.dataset.generated) {
        this.removeChild(child);
        i--;
      }
    }
  }

  renderTemplate(template: HTMLTemplateElement, context: object, slotId: string) {
    let templateContent = TemplateHelper.renderTemplate(template, context);

    let div = document.createElement('div');
    div.slot = slotId;
    div.dataset.generated = 'template';
    div.appendChild(templateContent);

    this.appendChild(div);

    this.fireCustomEvent('templateRendered', { context: context, element: div });

    return html`
      <slot name=${slotId}></slot>
    `;
  }
}
