import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { state } from '@microsoft/mgt-element/node_modules/lit-element/lib/decorators';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { debounce } from '../../utils/Utils';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { strings } from './strings';
import { styles } from './mgt-generic-picker-css';
import { DriveItem } from '@microsoft/microsoft-graph-types';

@customElement('mgt-generic-picker')
export class MgtGenericPicker extends MgtTemplatedComponent {
  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtGenericPicker
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled: boolean;

  /**
   *  array of user picked drive items.
   * @type {DriveItem[]}
   */
  @property({
    attribute: 'selected-drive-items',
    type: Array
  })
  public selectedDriveItems: DriveItem[];

  /**
   * array of entities to be searched on the graph
   *
   * @type {string[]}
   * @memberof MgtGenericPicker
   */
  @property({
    attribute: 'entity-types',
    converter: value => {
      return value.split(',').map(v => v.trim()?.toLowerCase());
    },
    type: String
  })
  public entityTypes: string[];

  // if search is still loading don't load "x not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  @state() private _isFocused = false;
  private _debouncedSearch: { (): void; (): void };
  private _driveIds: string[];

  protected get strings() {
    return strings;
  }

  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtGenericPicker
   */
  protected userInput: string;

  constructor() {
    super();
    this.clearState();
    this._showLoading = true;
    this.addEventListener('itemClick', e => this.handleItemSelect(e));
  }

  public render(): TemplateResult {
    const renderSelectedDriveItemsTpl = this.renderSelectedDriveItems(this.selectedDriveItems);
    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);

    const inputClasses = {
      focused: this._isFocused,
      'generic-picker': true,
      disabled: this.disabled
    };
    return html`
      <div
        dir=${this.direction}
        class=${classMap(inputClasses)}>
        <div class="selected-list">
          ${renderSelectedDriveItemsTpl}${flyoutTemplate}
        </div>
      </div>
      `;
  }

  protected renderInput(): TemplateResult {
    const hasSelections = !!this.selectedDriveItems.length;

    const placeholder = 'Enter some text';
    const inputClasses = {
      'search-box': true,
      'search-box-start': hasSelections
    };
    return html`
    <div class="${classMap(inputClasses)}">
      <input
        id="people-picker-input"
        class="search-box__input"
        type="text"
        placeholder=${placeholder}
        role="input"
        @keydown="${this.onUserKeyDown}"
        @keyup="${this.onUserKeyUp}"
        @blur=${this.lostFocus}
        @click=${this.handleFlyout}
        ?disabled=${this.disabled}
      />
    </div>
  `;
  }

  private handleFlyout() {
    // this._showLoading = true;
    this.showFlyout();
  }

  private gainedFocus() {
    this._isFocused = true;
    if (this.input) {
      this.input.focus();
    }
    this._showLoading = true;
    this.loadState();
  }

  private lostFocus() {
    this._isFocused = false;
    this.requestUpdate();
  }

  protected renderSelectedDriveItems(driveItems?: DriveItem[]): TemplateResult {
    if (!this.selectedDriveItems || !driveItems.length) {
      return null;
    }

    return html`<div class="selected-list__options">
      ${driveItems.slice(0, driveItems.length).map(item => {
        html`<div class="selected-list__file-wrapper">
          ${this.renderSelectedDriveItem(item)}

          <div class="selected-list__file-wrapper__overflow">
            <div class="selected-list__file-wrapper__overflow__gradient"></div>
            <div
              tabindex="0"
              aria-label="close-icon"
              class="selected-list__file-wrapper__overflow__close-icon"
              @click="${e => this.removeDriveItem(item, e)}"
            >
              \uE711
            </div>
          </div>
        </div>`;
      })}
    </div>`;
  }
  protected renderSelectedDriveItem(driveItem: DriveItem): TemplateResult {
    return html`
      <mgt-file view="twoLines" file-details=driveItem></mgt-file>
    `;
  }

  protected removeDriveItem(driveItem: DriveItem, event: Event): void {
    console.log('Removing selected driveItem ', driveItem);
  }
  /**
   * Focuses the input element when focus is called
   *
   * @param {FocusOptions} [options]
   * @memberof MgtPeoplePicker
   */
  public focus(options?: FocusOptions) {
    this.gainedFocus();
    if (!this.input) {
      return;
    }
    this.input.focus(options);
    this.input.select();
  }

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtGenericPicker
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * Gets the input element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtGenericPicker
   */
  protected get input(): HTMLInputElement {
    return this.renderRoot.querySelector('.search-box__input');
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtGenericPicker
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }

  /**
   * Render the flyout chrome.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtGenericPicker
   */
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    return html`
        <mgt-flyout light-dismiss class="flyout">
            ${anchor}
            <div slot="flyout" class="flyout-root" @wheel=${(e: WheelEvent) => this.handleSectionScroll(e)}>
            ${this.renderFlyoutContent()}
            </div>
        </mgt-flyout>
        `;
  }

  /**
   * Render the appropriate state in the results flyout.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtGenericPicker
   */
  protected renderFlyoutContent(): TemplateResult {
    if (this.isLoadingState || this._showLoading) {
      return this.renderLoading();
    }

    return this.renderSearchResults();
  }

  protected renderSearchResults() {
    const hasFiles = this.entityTypes.includes('file-list');
    let fileListTemplate: TemplateResult;

    if (hasFiles && this._driveIds && this._driveIds.length > 0) {
      // TODO: there's a slight lag here. Maybe make the loading happen
      // some more or use the mgt-file-list load time? Also check the
      // caching
      const driveIds = this._driveIds.join(',');
      fileListTemplate = html`
        <div>Files</div>
        <mgt-file-list
            hide-more-files-button
            files=${driveIds}
            insight-type="used"></mgt-file-list>
      `;
    } else {
      // TODO: page-size not supported when insight-type is available. Maybe
      // have the page-size support handled by MGT?
      fileListTemplate = html`
        <div>Files</div>
        <mgt-file-list hide-more-files-button insight-type="used"></mgt-file-list>`;
    }

    // TODO: how to render other entities i.e. people list
    const peopleListTemplate = this.renderPeople();

    const resultsTemplate = html`
    <div>
      ${peopleListTemplate}
      ${fileListTemplate ? fileListTemplate : ''}
    </div>
    `;
    return resultsTemplate;
  }

  protected renderPeople(): TemplateResult {
    return html`
    <div>
      <p>People</p>
      <ul>
        <li><mgt-person person-query="me" view="twoLines"></mgt-person></li>
        <li><mgt-person person-query="me" view="twoLines"></mgt-person></li>
        <li><mgt-person person-query="me" view="twoLines"></mgt-person></li>
        <li><mgt-person person-query="me" view="twoLines"></mgt-person></li>
      </ul>
    </div>
    `;
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtGenericPicker
   */
  protected renderLoading(): TemplateResult {
    return (
      this.renderTemplate('loading', null) ||
      html`
        <div class="message-parent">
        <mgt-spinner></mgt-spinner>
        <div label="loading-text" aria-label="loading" class="loading-text">
            ${this.strings.loadingMessage}
        </div>
        </div>`
    );
  }

  // stop propagating wheel event to flyout so mouse scrolling works
  private handleSectionScroll(e: WheelEvent) {
    const target = this.renderRoot.querySelector('.flyout-root') as HTMLElement;
    if (target) {
      if (
        !(e.deltaY < 0 && target.scrollTop === 0) &&
        !(e.deltaY > 0 && target.clientHeight + target.scrollTop >= target.scrollHeight - 1)
      ) {
        e.stopPropagation();
      }
    }
  }

  /**
   * Async query to Graph
   */
  protected async loadState(): Promise<void> {
    const input = this.userInput.toLowerCase();
    const provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);
      // TODO: check the required permissions are present
      // TODO: check the allowed entitytypes are set
      console.log('input ', !input.length && this._isFocused);
      if (!input.length && this._isFocused) {
        // TODO: figure out what to load here? Check the mgt-file-list default behavior
        console.log('We i here');
        this._showLoading = false;
        this._isFocused = false;
        return;
      }
      this._showLoading = false;
      if (input) {
        // const fileList = await graph.api('/beta/search/query').version('beta');
        // TODO: fetch files/driveItems that fulfill the search query and extract their IDs
        const fileList = [
          '01BYE5RZ2XNDO2A5W3IRCYJYDR3QOECFQC',
          '01BYE5RZ2TWC4BCH4FQJB3HMWWDKXBHK2S',
          '01BYE5RZ5ZEXZD2H26FRALRJATDOA2TQGJ'
        ];
        this._driveIds = fileList;
      }
    }
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: KeyboardEvent): void {
    this.flyout.close();
    const input = event.target as HTMLInputElement;
    this.userInput = input.value;
    console.log({ keyDownText: this.userInput });
  }

  /**
   * Adds debounce method for set delay on user input
   */
  private onUserKeyUp(event: KeyboardEvent): void {
    const input = event.target as HTMLInputElement;
    this.userInput = input.value;
    // this.gainedFocus();
    // TODO: Bug: Figure out when to set this._isFocused = true
    this.flyout.close();
    this.handleUserSearch();
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtGenericPicker
   */
  protected clearState(): void {
    this._clearInput();
    this.selectedDriveItems = [];
  }

  /**
   * Clear the typed text from the input element
   */
  private _clearInput() {
    this.userInput = '';
  }

  /**
   * Tracks event on user input in search
   * @param input - input text
   */
  private handleUserSearch() {
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        const loadingTimeout = setTimeout(() => {
          this._showLoading = true;
        }, 50);

        await this.loadState();
        clearTimeout(loadingTimeout);
        this._showLoading = false;
        this.showFlyout();

        // this._arrowSelectionCount = 0;
      }, 400);
    }

    this._debouncedSearch();
  }

  protected handleItemSelect(event: Event): void {
    // this.handleItemSelect(item, event);
    // console.log('Clicked an item ', item);
    this.fireCustomEvent('itemClick');

    console.log('Click item event ', event);
  }

  /**
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected requestStateUpdate(force?: boolean) {
    if (force) {
      this.selectedDriveItems = [];
    }

    return super.requestStateUpdate(force);
  }
}
