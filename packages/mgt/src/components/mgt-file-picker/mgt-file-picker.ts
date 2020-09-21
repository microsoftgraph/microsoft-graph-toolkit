import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { getMyInsights, InsightsItem } from './graph.files';
import { styles } from './mgt-file-picker-css';

/**
 * The various insights endpoints used to pull data from.
 *
 * @export
 * @enum {number}
 */
export enum InsightsDataSource {
  trending,
  shared,
  used
}

// The default data source.
const DEFAULT_DATA_SOURCE = InsightsDataSource.trending;

/**
 * foo
 *
 * @export
 * @class MgtFilePicker
 * @extends {MgtTemplatedComponent}
 */
@customElement('mgt-file-picker')
export class MgtFilePicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  @property({
    attribute: 'data-source',
    converter: value => {
      value = value.toLowerCase();
      if (typeof InsightsDataSource[value] === 'undefined') {
        return DEFAULT_DATA_SOURCE;
      } else {
        return InsightsDataSource[value];
      }
    }
  })
  public dataSource: InsightsDataSource;

  private _insights: InsightsItem[];

  constructor() {
    super();

    this.dataSource = DEFAULT_DATA_SOURCE;
  }

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtPerson
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * Render the component
   *
   * @returns
   * @memberof MgtFilePicker
   */
  render() {
    const root = html`
      <div class="root">
        ${this.renderButton()}
      </div>
    `;

    const flyoutContent = html`
      <div slot="flyout">
        ${this.renderFlyoutContent()}
      </div>
    `;

    return html`
      <mgt-flyout light-dismiss class="flyout">
        ${root} ${flyoutContent}
      </mgt-flyout>
    `;
  }

  /**
   * Render the button used to invoke the flyout
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderButton(): TemplateResult {
    return html`
      <div class="button" @click=${() => this.toggleFlyout()}>
        Choose from OneDrive
      </div>
    `;
  }

  /**
   * Render the contents of the flyout, visible when open
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderFlyoutContent(): TemplateResult {
    return html`
      ${this._insights ? this._insights.length : 0} Files
    `;
  }

  /**
   * Toggle the flyout visiblity
   *
   * @protected
   * @memberof MgtFilePicker
   */
  protected toggleFlyout(): void {
    if (this.flyout.isOpen) {
      this.flyout.close();
    } else {
      this.flyout.open();
    }
  }

  protected async loadState() {
    // Check if signed in.
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    const graph = provider.graph;
    this._insights = await getMyInsights(graph, this.dataSource);
  }
}
