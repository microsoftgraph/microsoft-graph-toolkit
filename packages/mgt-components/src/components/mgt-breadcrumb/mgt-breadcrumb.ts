import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { repeat } from 'lit/directives/repeat.js';
import { fluentBreadcrumb, fluentBreadcrumbItem } from '@fluentui/web-components';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { customElement, MgtBaseComponent } from '@microsoft/mgt-element';
import { styles } from './mgt-breadcrumb-css';

registerFluentComponents(fluentBreadcrumb, fluentBreadcrumbItem);

/**
 * Defines a base type for breadcrumb data
 */
export type BreadcrumbInfo = {
  /**
   * unique identifier for the breadcrumb
   *
   * @type {string}
   */
  id: string;
  /**
   * Name of the breadcrumb, used to display the item.
   *
   * @type {string}
   */
  name: string;
};

/**
 * Custom breadcrumb component
 *
 * @fires {CustomEvent<BreadcrumbInfo>} breadcrumbclick - Fired when a breadcrumb is clicked. Will not fire when the last breadcrumb is clicked.
 *
 * @cssprop --breadcrumb-base-font-size - {Length} Breadcrumb base font size. Default is 18px.
 *
 * @export
 * @class MgtBreadcrumb
 * @extends {MgtBaseComponent}
 */
@customElement('breadcrumb')
export class MgtBreadcrumb extends MgtBaseComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }
  /**
   * An array of nodes to show in the breadcrumb
   *
   * @type {BreadcrumbInfo[]}
   * @readonly
   * @memberof MgtFileList
   */
  @property()
  private breadcrumb: BreadcrumbInfo[] = [];

  private isLastCrumb = (b: BreadcrumbInfo): boolean => this.breadcrumb.indexOf(b) === this.breadcrumb.length - 1;

  /**
   * Renders the component
   *
   * @return {*}  {TemplateResult}
   * @memberof MgtBreadcrumb
   */
  public render(): TemplateResult {
    return html`
      <fluent-breadcrumb class="breadcrumb">
        ${repeat(
          this.breadcrumb,
          b => b.id,
          b =>
            html`
              <fluent-breadcrumb-item
                class="breadcrumb-item"
                @click=${() => this.handleBreadcrumbClick(b)}
                @keypress=${(e: KeyboardEvent) => this.handleBreadcrumbKeyPress(e, b)}
              >
                <span
                  tabindex=${this.isLastCrumb(b) ? -1 : 0}
                  class=${this.isLastCrumb(b) ? '' : 'interactive-breadcrumb'}
                >
                  ${b.name}
                </span>
              </fluent-breadcrumb-item>
            `
        )}
      </fluent-breadcrumb>
    `;
  }

  private handleBreadcrumbKeyPress(event: KeyboardEvent, b: BreadcrumbInfo): void {
    if (event.key === 'Enter' || event.key === ' ') {
      this.handleBreadcrumbClick(b);
    }
  }
  private handleBreadcrumbClick(b: BreadcrumbInfo): void {
    // last crumb does nothing
    if (this.isLastCrumb(b)) return;
    this.fireCustomEvent('breadcrumbclick', b);
  }
}
