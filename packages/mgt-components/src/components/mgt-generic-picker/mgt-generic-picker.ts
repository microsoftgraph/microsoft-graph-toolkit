import { MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { state } from '@microsoft/mgt-element/node_modules/lit-element/lib/decorators';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { debounce } from '../../utils/Utils';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { strings } from './strings';
import { styles } from './mgt-generic-picker-css';

@customElement('mgt-generic-picker')
export class MgtGenericPicker extends MgtTemplatedComponent {
  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled: boolean;
  // if search is still loading don't load "x not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  @state() private _isFocused = false;
  private _debouncedSearch: { (): void; (): void };
  private _fileIds: string[];

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
   * @memberof MgtPeoplePicker
   */
  protected userInput: string;

  constructor() {
    super();
    this.clearState();
    this._showLoading = true;
  }

  public render(): TemplateResult {
    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);
    return html`
      ${flyoutTemplate}
      `;
  }

  protected renderInput(): TemplateResult {
    const placeholder = 'Enter some text';
    const inputClasses = {
      focused: this._isFocused,
      'people-picker': true,
      disabled: this.disabled
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

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * Gets the input element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  protected get input(): HTMLInputElement {
    return this.renderRoot.querySelector('.search-box__input');
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
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
   * @memberof MgtPeoplePicker
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
   * @memberof MgtPeoplePicker
   */
  protected renderFlyoutContent(): TemplateResult {
    if (this.isLoadingState || this._showLoading) {
      return this.renderLoading();
    }

    return this.renderSearchResults();
  }

  protected renderSearchResults() {
    if (this._fileIds && this._fileIds.length > 0) {
      // TODO: there's a slight lag here. Maybe make the loading happen
      // some more or use the mgt-file-list load time? Also check the
      // caching
      return html`
        <div>Files</div>
        <mgt-file-list
            hide-more-files-button
            files=${this._fileIds.join(',')}></mgt-file-list>
      `;
    }
    return html`
        <div>Files</div>
        <mgt-file-list hide-more-files-button></mgt-file-list>`;
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtPeoplePicker
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
      this._showLoading = true;
      // TODO: check the required permissions are present
      // TODO: check the allowed entitytypes are set
      console.log('Loadstate is focused ', this._isFocused);
      if (!input.length && this._isFocused) {
        // TODO: figure out what to load here? Check the mgt-file-list default behavior
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
        this._fileIds = fileList;
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
    this.flyout.close();
    this.handleUserSearch();
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected clearState(): void {
    this.userInput = '';
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
}
