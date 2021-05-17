/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import '../../styles/style-helper';
import '../sub-components/mgt-spinner/mgt-spinner';
import { debounce } from '../../utils/Utils';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { styles } from './mgt-search-suggestion-css';

import { strings } from './strings';
import {
  getSuggestions,
  SuggestionFile,
  SuggestionPeople,
  Suggestions,
  SuggestionText
} from '../../graph/graph.suggestions';
import { IDynamicPerson } from '../../graph/types';

/**
 * An interface used to mark an object as 'focused',
 * so it can be rendered differently.
 *
 * @interface IFocusable
 */
interface IFocusable {
  // tslint:disable-next-line: completed-docs
  isFocused: boolean;
}

/**
 * Web component used to search for people from the Microsoft Graph
 *
 * @export
 * @cssprop --suggestion-item-background-color--hover - {Color} background color for an hover item
 * @cssprop --suggestion-list-background-color - {Color} background color
 * @cssprop --suggestion-list-text-color - {Color} Text Suggestion font color
 *
 */
@customElement('mgt-search-suggestion')
export class MgtSearchSuggestion extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  protected get strings() {
    return strings;
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
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({ attribute: 'value' })
  public get value() {
    return this.input.value;
  }

  public set value(_value) {
    this.input.value = _value;
  }

  private _testValue = 'good';

  @property({ attribute: 'suggestion-value' })
  public get suggestionValue() {
    return this.getFocusItemValue();
  }

  public testExpose() {
    return 'expose';
  }

  /**
   * Placeholder text.
   *
   * @type {string}
   * @memberof MgtSearchSuggestion
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Determines whether component should be disabled or not
   *
   * @type {boolean}
   * @memberof MgtSearchSuggestion
   */
  @property({
    attribute: 'disabled',
    type: Boolean
  })
  public disabled: boolean;

  @property({
    attribute: 'people-suggestion-value',
    type: String
  })
  private _peopleSuggestionValue: string;

  public set peopleSuggestionValue(peopleSuggestionValue: string) {
    this._peopleSuggestionValue = peopleSuggestionValue;
  }

  public get peopleSuggestionValue() {
    return this._peopleSuggestionValue;
  }

  @property({
    attribute: 'file-suggestion-value',
    type: String
  })
  private _fileSuggestionValue: string;

  public set fileSuggestionValue(fileSuggestionValue: string) {
    this._fileSuggestionValue = fileSuggestionValue;
  }

  public get fileSuggestionValue() {
    return this._fileSuggestionValue;
  }

  @property({
    attribute: 'text-suggestion-value',
    type: String
  })
  private _textSuggestionValue: string;

  public set textSuggestionValue(textSuggestionValue: string) {
    this._textSuggestionValue = textSuggestionValue;
  }

  public get textSuggestionValue() {
    return this._textSuggestionValue;
  }

  @property({
    attribute: 'max-file-suggestion-count',
    type: Number
  })
  private _maxFileSuggestionCount: number = 3;

  public set maxFileSuggestionCount(maxFileSuggestionCount: number) {
    this._maxFileSuggestionCount = maxFileSuggestionCount;
  }

  public get maxFileSuggestionCount() {
    return this._maxFileSuggestionCount;
  }

  @property({
    attribute: 'max-text-suggestion-count',
    type: Number
  })
  private _maxTextSuggestionCount: number = 3;

  public set maxTextSuggestionCount(maxTextSuggestionCount: number) {
    this._maxTextSuggestionCount = maxTextSuggestionCount;
  }

  public get maxTextSuggestionCount() {
    return this._maxTextSuggestionCount;
  }

  @property({
    attribute: 'max-people-suggestion-count',
    type: Number
  })
  private _maxPeopleSuggestionCount: number = 3;

  public set maxPeopleSuggestionCount(maxPeopleSuggestionCount: number) {
    this._maxPeopleSuggestionCount = maxPeopleSuggestionCount;
  }

  public get maxPeopleSuggestionCount() {
    return this._maxPeopleSuggestionCount;
  }

  @property({
    attribute: 'selected-entity-types',
    type: String
  })
  private _selectedEntityTypes: string = 'file, text, people';

  public set selectedEntityTypes(selectedEntityTypes: string) {
    this._selectedEntityTypes = selectedEntityTypes;
  }

  @property({
    attribute: 'cvid',
    type: String
  })
  private _cvid: string = 'd8e48cff-9cac-40c9-0b5c-e6d24488f781';

  public set cvid(cvid: string) {
    this._cvid = cvid;
  }

  @property({
    attribute: 'text-decorations',
    type: String
  })
  private _textDecorations: string = 'd8e48cff-9cac-40c9-0b5c-e6d24488f781';

  public set textDecorations(textDecorations: string) {
    this._textDecorations = textDecorations;
  }

  public get selectedEntityTypes() {
    return this._selectedEntityTypes;
  }

  public onEnterKeyPressCallback: (originalInput, suggestionValue) => {};

  public onClickCallback: (suggestionValue) => {};

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtSearchSuggestion
   */
  protected userInput: string;

  // if search is still loading don't load "people not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  // tracking of user arrow key input for selection
  private _arrowSelectionCount: number = 0;

  private _debouncedSearch: { (): void; (): void };

  private arrow_initial = false;

  @internalProperty() private _isFocused = false;

  @internalProperty() private _foundPeopleSuggestion: SuggestionPeople[];

  @internalProperty() private _foundTextSuggestion: SuggestionText[];

  @internalProperty() private _foundFileSuggestion: SuggestionFile[];

  @internalProperty() private _foundSuggestion: Suggestions;

  @internalProperty() private _suggestionItemsMap: Map<string, any> = new Map();

  constructor() {
    super();
    this.clearState();
    this._showLoading = true;
    this.disabled = false;
  }

  /**
   * Get the scopes required for search suggestion
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtSearchSuggestion
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set(['']) // awaiting on search suggestion onboard
    ];
  }

  /**
   * Focuses the input element when focus is called
   *
   * @param {FocusOptions} [options]
   * @memberof MgtSearchSuggestion
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
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns {TemplateResult}
   * @memberof MgtSearchSuggestion
   */
  public render(): TemplateResult {
    const defaultTemplate = this.renderTemplate('default', { people: this._foundSuggestion });
    if (defaultTemplate) {
      return defaultTemplate;
    }

    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);

    const inputClasses = {
      focused: this._isFocused,
      'search-suggestion': true,
      disabled: this.disabled
    };

    return html`
      <div dir=${this.direction} class=${classMap(inputClasses)} @click=${e => this.focus(e)}>
        <div class="selected-list">
          ${flyoutTemplate}
        </div>
      </div>
    `;
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtSearchSuggestion
   */
  protected clearState(): void {
    this.userInput = '';
  }

  /**
   * Request to reload the state.
   * Use reload instead of load to ensure loading events are fired.
   *
   * @protected
   * @memberof MgtBaseComponent
   */
  protected requestStateUpdate(force?: boolean) {
    return super.requestStateUpdate(force);
  }

  /**
   * Render the input text box.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearchSuggestion
   */
  protected renderInput(): TemplateResult {
    const inputClasses = {
      'search-box': true
    };

    return (
      this.renderTemplate('search-suggestion-input', null) ||
      html`
      <div class="${classMap(inputClasses)}">
        <input
          id="search-suggestion-input"
          class="search-box__input"
          type="text"
          placeholder="Search sth..."
          label="search-suggestion-input"
          aria-label="search-suggestion-input"
          role="input"
          @keydown="${this.onUserKeyDown}"
          @keyup="${this.onUserKeyUp}"
          @blur=${this.lostFocus}
          @click=${this.handleFlyout}
          ?disabled=${this.disabled}
        />
      </div>
    `
    );
  }

  /**
   * Render the flyout chrome.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearchSuggestion
   */
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    return html`
      <mgt-flyout light-dismiss class="flyout">
        ${anchor}
        <div slot="flyout" class="flyout-root">
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
   * @memberof MgtSearchSuggestion
   */
  protected renderFlyoutContent(): TemplateResult {
    if (this.isLoadingState || this._showLoading) {
      return this.renderLoading();
    }

    let fileSuggestions = this._foundFileSuggestion;
    let textSuggestions = this._foundTextSuggestion;
    let peopleSuggestions = this._foundPeopleSuggestion;

    if (
      (!peopleSuggestions || peopleSuggestions.length === 0 || this._maxPeopleSuggestionCount === 0) &&
      (!textSuggestions || textSuggestions.length === 0 || this._maxTextSuggestionCount === 0) &&
      (!fileSuggestions || fileSuggestions.length === 0 || this._maxFileSuggestionCount === 0)
    ) {
      return this.renderNoData();
    }

    return html`
      <div class="suggestion-container">
        ${this.renderTextSearchResults(textSuggestions)} ${this.renderPeopleSearchResults(peopleSuggestions)}
        ${this.renderFileSearchResults(fileSuggestions)}
      </div>
    `;

    // render text result
    // render file
  }

  /**
   * Render the loading state.
   *
   * @protected
   * @returns
   * @memberof MgtSearchSuggestion
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
        </div>
      `
    );
  }

  /**
   * Render the state when no results are found for the search query.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtSearchSuggestion
   */
  protected renderNoData(): TemplateResult {
    return (
      this.renderTemplate('error', null) ||
      this.renderTemplate('no-data', null) ||
      html`
        <div class="message-parent">
          <div label="search-error-text" aria-label="We didn't find any matches." class="search-error-text">
            ${this.strings.noResultsFound}
          </div>
        </div>
      `
    );
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    const input = this.userInput;
    if (provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);
      if (this._isFocused) {
        var SuggestionQueryConfig = {
          maxFileCount: this._maxFileSuggestionCount,
          maxPeopleCount: this._maxPeopleSuggestionCount,
          maxTextCount: this._maxTextSuggestionCount,
          queryString: input,
          queryEntities: this._selectedEntityTypes,
          cvid: this._cvid,
          textDecorations: this._textDecorations
        };
        var suggestions = await getSuggestions(graph, SuggestionQueryConfig);
        this._foundSuggestion = suggestions;
        this._foundTextSuggestion = this.filterTextSuggestions(suggestions.textSuggestions);
        this._foundFileSuggestion = this.filterFileSuggestions(suggestions.fileSuggestions);
        this._foundPeopleSuggestion = this.filterPeopleSuggestions(suggestions.peopleSuggestions);
      }
      this._showLoading = false;
      this.clearArrowSelection();
    }
  }

  /**
   * Hide the results flyout.
   *
   * @protected
   * @memberof MgtSearchSuggestion
   */
  protected hideFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtSearchSuggestion
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }

  private clearInput() {
    this.input.value = '';
    this.userInput = '';
  }

  private handleFlyout() {
    // handles hiding control if default people have no more selections available
    let shouldShow = true;
    if (shouldShow) {
      window.requestAnimationFrame(() => {
        // Mouse is focused on input
        this.showFlyout();
      });
    }
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
   * Adds debounce method for set delay on user input
   */
  private onUserKeyUp(event: KeyboardEvent): void {
    if (event.keyCode === 40 || event.keyCode === 39 || event.keyCode === 38 || event.keyCode === 37) {
      // keyCodes capture: down arrow (40), right arrow (39), up arrow (38) and left arrow (37)
      return;
    }

    const input = event.target as HTMLInputElement;

    if (event.code === 'Escape') {
      this.clearInput();
    } else {
      this.userInput = input.value;
      this.handleUserSearch();
    }
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

        this._arrowSelectionCount = 0;
      }, 400);
    }

    this._debouncedSearch();
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: KeyboardEvent): void {
    console.log('Press key:', event.keyCode);
    if (!this.flyout.isOpen) {
      return;
    }
    if (event.keyCode === 40 || event.keyCode === 38 || event.keyCode === 9) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this.input.value.length > 0) {
        event.preventDefault();
      }
    }
    if (event.keyCode === 13) {
      //  and enter (13)
      this.onEnterKeyPressCallback(this.input.value, this.getFocusItemValue());
      this.hideFlyout();
      (event.target as HTMLInputElement).value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: KeyboardEvent): void {
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    if (peopleList && peopleList.length) {
      // update arrow count
      if (event.keyCode === 38) {
        // up arrow
        if (this.arrow_initial) {
          this._arrowSelectionCount = (this._arrowSelectionCount - 1 + peopleList.length) % peopleList.length;
        } else {
          this.arrow_initial = true;
          this._arrowSelectionCount = peopleList.length - 1;
          this._arrowSelectionCount = 0;
        }
      }
      if (event.keyCode === 40 || event.keyCode === 9) {
        // down arrow or tab
        if (this.arrow_initial) {
          this._arrowSelectionCount = (this._arrowSelectionCount + 1) % peopleList.length;
        } else {
          this.arrow_initial = true;
          this._arrowSelectionCount = 0;
        }
      }

      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.length; i++) {
        peopleList[i].classList.remove('suggestion-focused');
      }

      // set selected background
      const focusedItem = peopleList[this._arrowSelectionCount];
      if (focusedItem) {
        focusedItem.classList.add('suggestion-focused');
        focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'start' });
      }
    }
  }

  private clearArrowSelection(): void {
    this.arrow_initial = false;
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    for (let i = 0; i < peopleList.length; i++) {
      peopleList[i].classList.remove('suggestion-focused');
    }
  }

  private getFocusItemValue() {
    const peopleList = this.renderRoot.querySelectorAll('.suggestion-common-container');
    for (let i = 0; i < peopleList.length; i++) {
      if (peopleList[i].classList.contains('suggestion-focused')) {
        return this._suggestionItemsMap.get(peopleList[i].id);
      }
    }
  }

  /**
   * Filters people searched from already selected people
   * @param people - array of people returned from query to Graph
   */
  private filterPeopleSuggestions(people: SuggestionPeople[]): SuggestionPeople[] {
    for (let person of people) {
      this._suggestionItemsMap.set(person.referenceId, person);
    }
    return people;
  }

  private filterFileSuggestions(files: SuggestionFile[]): SuggestionFile[] {
    for (let file of files) {
      this._suggestionItemsMap.set(file.referenceId, file);
    }
    return files;
  }

  private filterTextSuggestions(texts: SuggestionText[]): SuggestionText[] {
    // check if people need to be updated
    // ensuring people list is displayed
    // find ids from selected people
    for (let text of texts) {
      this._suggestionItemsMap.set(text.referenceId, text);
    }
    return texts;
  }

  // render People search result
  protected renderPeopleSearchResults(people?: SuggestionPeople[]) {
    people = people || this._foundPeopleSuggestion;
    if (people.length < 1) return html``;
    const input = this.userInput;
    return html`
        ${this.renderSuggestionEntityLabelPeople()}
        ${this.renderSuggestionEntityPeople(people)}
    `;
  }

  // render File search result
  protected renderFileSearchResults(files?: SuggestionFile[]) {
    files = files || this._foundFileSuggestion;
    if (files.length < 1) return html``;
    const input = this.userInput;
    return html`
        ${this.renderSuggestionEntityLabelFile()}
        ${this.renderSuggestionEntityFile(files)}
    `;
  }

  // render Text search result
  protected renderTextSearchResults(texts?: SuggestionText[]) {
    texts = texts || this._foundTextSuggestion;
    if (texts.length < 1) return html``;
    const input = this.userInput;
    return html`
        ${this.renderSuggestionEntityLabelText()}
        ${this.renderSuggestionEntityText(texts)}
    `;
  }

  // Template Rendering

  protected renderSuggestionEntityLabelPeople() {
    return (
      this.renderTemplate('search-suggestion-label-people', null) ||
      html`
      <div class="suggestion-entity-label">
        People
      </div>
      `
    );
  }

  protected renderSuggestionEntityLabelText() {
    return (
      this.renderTemplate('search-suggestion-label-text', {}) ||
      html`
      <div class="suggestion-entity-label">
        Suggested Searches
      </div>
      `
    );
  }

  protected renderSuggestionEntityLabelFile() {
    return (
      this.renderTemplate('search-suggestion-label-file', {}) ||
      html`
      <div class="suggestion-entity-label">
        Files
      </div>
      `
    );
  }

  protected renderSuggestionEntityText(texts?: SuggestionText[]) {
    texts = texts || this._foundTextSuggestion;
    if (texts.length < 1) return html``;
    const input = this.userInput;
    return (
      this.renderTemplate('search-suggestion-text', { texts }) ||
      html`
      ${repeat(texts, text => {
        return html`
          <div
            class="suggestion-text-container suggestion-common-container"
            id="${text.referenceId}"
            @click="${e => this.onClickCallback(text)}"
          >
            <div class="suggestion-content-container">
              <div class="suggestion-text-description">
                <b>${text.text.slice(0, input.length)}</b><span>${text.text.slice(input.length)}</span>
              </div>
            </div>
          </div>
        `;
      })}
    `
    );
  }

  protected renderSuggestionEntityPeople(people?: SuggestionPeople[]) {
    return (
      this.renderTemplate('search-suggestion-people', { people }) ||
      html`
      ${repeat(people, person => {
        return html`
        <div
            class="suggestion-common-container"
            id="${person.referenceId}"
            @click="${e => {
              this.onClickCallback(person);
            }}"
          >
          <mgt-person
            .personDetails=${this.ConvertSuggestionPersonToIDynamicPerson(person)}
            .fetchImage=${true}
            class="mgt-search-suggestion-person-default"
            view="threeLines"
          ></mgt-person>

        </div>
        `;
      })}
    `
    );
  }

  protected renderSuggestionEntityFile(files?: SuggestionFile[]) {
    return (
      this.renderTemplate('search-suggestion-file', { files }) ||
      html`
      ${repeat(files, file => {
        return html`
          <div
            class="suggestion-common-container"
            id="${file.referenceId}"
            @click="${e => {
              this.onClickCallback(file);
            }}"
          >


          <mgt-file
            .fileDetails=${this.ConvertSuggestionFileToDriveItem(file)}
            class="mgt-search-suggestion-file-default"
          ></mgt-file>


          </div>
        `;
      })}
    `
    );
  }

  // Covert Suggestion Entity to File/Text/People

  private ConvertSuggestionFileToDriveItem(file: SuggestionFile): DriveItem {
    let driveItem: DriveItem = {
      size: file.FileSize,
      webDavUrl: file.AccessUrl,
      fileSystemInfo: {
        lastModifiedDateTime: file.DateModified
      },
      name: file.name
    };
    return driveItem;
  }

  private ConvertSuggestionPersonToIDynamicPerson(person: SuggestionPeople): IDynamicPerson {
    let dynamicPerson: IDynamicPerson = {
      displayName: person.displayName,
      jobTitle: person.jobTitle,
      personImage: person.personImage,
      imAddress: person.imAddress
    };
    return dynamicPerson;
  }
}
