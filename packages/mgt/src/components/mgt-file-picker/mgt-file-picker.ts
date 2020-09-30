import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { getMyInsights, InsightsItem } from './graph.files';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { styles } from './mgt-file-picker-css';
import { getRelativeDisplayDate } from '../../utils/Utils';
import '../sub-components/mgt-spinner/mgt-spinner';

const strings = {
  buttonLabel: 'Pick from OneDrive',
  itemModifiedFormat: 'Modified {0}',
  seeAllItems: 'See all',
  resultsTitle: 'Recent'
};

const formatString = function(format: string, ...values: string[]): string {
  values.forEach((v, i) => (format = format.replace(`{${i}}`, v)));
  return format;
};

const isSignedIn = () => Providers.globalProvider && Providers.globalProvider.state === ProviderState.SignedIn;

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

  /**
   * The items array rendered by the control.
   *
   * @readonly
   * @memberof MgtFilePicker
   */
  public get items() {
    return this._items;
  }

  /**
   * The selected item
   *
   * @readonly
   * @memberof MgtFilePicker
   */
  public get selectedItem() {
    return this._selectedItem;
  }

  @internalProperty()
  private _selectedItem: InsightsItem;

  @internalProperty()
  private _items: InsightsItem[];

  private _doLoad: boolean;

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

  constructor() {
    super();

    this._items = null;
    this._selectedItem = null;
    this._doLoad = false;
  }

  /**
   * Render the component
   *
   * @returns
   * @memberof MgtFilePicker
   */
  render() {
    const root = html`
      <div class="root">${this.renderButton()}</div>
    `;

    const flyoutContent = html`
      <div slot="flyout">${this.renderFlyoutContent()}</div>
    `;

    const flyoutClasses = classMap({
      flyout: true,
      disabled: !isSignedIn(),
      loading: this.isLoadingState
    });

    return html`
      <mgt-flyout class=${flyoutClasses} light-dismiss>
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
    const buttonText = this._selectedItem ? this._selectedItem.resourceVisualization.title : strings.buttonLabel;

    return html`
      <div class="button" @click=${() => this.toggleFlyout()}>
        <div class="button__text">${buttonText}</div>
        <div class="button__icon">${getSvg(SvgIcon.ExpandDown)}</div>
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
    let contentTemplate = this.isLoadingState
      ? this.renderLoading()
      : html`
          <div class="items">
            ${repeat(this._items || [], i => i.id, i => this.renderItem(i))}
          </div>
        `;

    return html`
      <div class="flyout-root">
        <div class="header">
          <div class="header__title">${strings.resultsTitle}</div>
          <div class="header__all-items" @click=${e => this.handleAllItemsClick(e)}>
            ${strings.seeAllItems}
          </div>
        </div>
        ${contentTemplate}
      </div>
    `;
  }

  /**
   * Render the loading state of the results flyout
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderLoading(): TemplateResult {
    return html`
      <div class="spinner">
        <mgt-spinner></mgt-spinner>
      </div>
    `;
  }

  /**
   * Render an individual item
   *
   * @protected
   * @param {InsightsItem} item
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderItem(item: InsightsItem): TemplateResult {
    let lastUsedTemplate: TemplateResult = null;
    const lastUsed = item.lastUsed;
    if (lastUsed && lastUsed.lastModifiedDateTime) {
      const lastUsedDate = new Date(item.lastUsed.lastModifiedDateTime);
      const lastUsedStringFormat = strings.itemModifiedFormat;
      const relativeDateString = getRelativeDisplayDate(lastUsedDate);
      const lastUsedString = formatString(lastUsedStringFormat, relativeDateString);

      lastUsedTemplate = html`
        <div class="item__last-used">
          ${lastUsedString}
        </div>
      `;
    }

    const itemClasses = classMap({
      item: true,
      'item--selected': this._selectedItem && item.id === this._selectedItem.id
    });

    return html`
      <div class=${itemClasses} @click=${e => this.handleItemClick(item, e)}>
        <div class="item__icon">
          ${getSvg(SvgIcon.File)}
        </div>
        <div class="item__details">
          <div class="item__title" alt=${item.resourceVisualization.title}>
            ${item.resourceVisualization.title}
          </div>
          ${lastUsedTemplate}
        </div>
      </div>
    `;
  }

  /**
   * Toggle the flyout visiblity
   *
   * @protected
   * @memberof MgtFilePicker
   */
  protected toggleFlyout(): void {
    if (!isSignedIn()) {
      return;
    }

    if (this.flyout.isOpen) {
      this.flyout.close();
    } else {
      // Lazy load
      if (!this._doLoad) {
        this._doLoad = true;
        this.requestStateUpdate();
      }

      this.flyout.open();
    }
  }

  /**
   * Handle the click event on an item.
   *
   * @protected
   * @param {InsightsItem} item
   * @memberof MgtFilePicker
   */
  protected handleItemClick(item: InsightsItem, event: PointerEvent): void {
    if (this._selectedItem === item) {
      this._selectedItem = null;
    } else {
      this._selectedItem = item;
    }
  }

  /**
   * Handle the event when the user clicks to see all items.
   *
   * @protected
   * @param {PointerEvent} e
   * @memberof MgtFilePicker
   */
  protected handleAllItemsClick(e: PointerEvent): void {
    this.openFullPicker();
  }

  /**
   * Open the full OneDrive File Picker
   *
   * @protected
   * @memberof MgtFilePicker
   */
  protected openFullPicker(): void {
    console.log('Full picker.');
  }

  /**
   * Load the state of the control.
   *
   * @protected
   * @returns
   * @memberof MgtFilePicker
   */
  protected async loadState() {
    // Check if signed in
    if (!isSignedIn()) {
      return;
    }

    // Don't load by default.
    // Wait for the flyout to be opened first.
    if (!this._doLoad) {
      return;
    }

    // Artificial loading time
    //const delay = ms => new Promise(c => setTimeout(c, ms));
    //await delay(1000 * 5);

    const graph = Providers.globalProvider.graph;
    this._items = await getMyInsights(graph);

    // Reset the selected item if it doesn't match any of the new results.
    if (this._selectedItem && this._items.findIndex(v => v.id === this._selectedItem.id) === -1) {
      this._selectedItem = null;
    }
  }
}
