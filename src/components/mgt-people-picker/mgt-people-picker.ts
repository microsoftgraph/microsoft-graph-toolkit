/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, TemplateResult } from 'lit-element';
import { repeat } from 'lit-html/directives/repeat';
import { findPerson, getPeopleFromGroup } from '../../graph/graph.people';
import { getUser } from '../../graph/graph.user';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { debounce } from '../../utils/Utils';
import { IDynamicPerson } from '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-picker-css';

/**
 * Web component used to search for people from the Microsoft Graph
 *
 * @export
 * @class MgtPicker
 * @extends {MgtTemplatedComponent}
 *
 * @cssprop --people-list-background-color - {Color} People list background color
 * @cssprop --accent-color - {Color} Accent color
 */
@customElement('mgt-people-picker')
export class MgtPeoplePicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * user the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * containing object of MgtPersonDetails.
   * @type {MgtPersonDetails}
   */
  @property({
    attribute: 'people',
    type: Object
  })
  public people: IDynamicPerson[];

  /**
   * determining how many people to show in list.
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public groupId: string;

  /**
   *  array of user picked people.
   * @type {Array<any>}
   */
  @property({
    attribute: 'selected-people',
    type: Array
  })
  public selectedPeople: IDynamicPerson[];

  // if search is still loading don't load "people not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  /**
   * determines if login menu popup should be showing
   * @type {boolean}
   */
  @property({ attribute: false }) private _showFlyout: boolean;

  // User input in search
  private _userInput: string;
  private _groupId: string;
  // tracking of user arrow key input for selection
  private arrowSelectionCount: number = 0;
  // List of people requested if group property is provided
  private groupPeople: any[];
  private debouncedSearch: { (): void; (): void };

  constructor() {
    super();

    this._showLoading = true;
    this._showFlyout = false;
    this._groupId = null;
    this._userInput = '';
    this.showMax = 6;
    this.selectedPeople = [];
    this.handleWindowClick = this.handleWindowClick.bind(this);
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtLogin
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('click', this.handleWindowClick);
  }

  /**
   * Invoked each time the custom element is disconnected from the document's DOM
   *
   * @memberof MgtLogin
   */
  public disconnectedCallback() {
    window.removeEventListener('click', this.handleWindowClick);
    super.disconnectedCallback();
  }

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtPersonCard
   */
  public attributeChangedCallback(att: string, oldval: string, newval: string) {
    super.attributeChangedCallback(att, oldval, newval);

    if (att === 'group-id' && oldval !== newval) {
      this.requestStateUpdate();
    }
  }

  /**
   * Focuses the input element when focus is called
   *
   * @param {FocusOptions} [options]
   * @memberof MgtPeoplePicker
   */
  public focus(options?: FocusOptions) {
    const peopleInput = this.renderRoot.querySelector('.people-selected-input') as HTMLInputElement;
    if (!peopleInput) {
      return;
    }

    peopleInput.focus(options);
    peopleInput.select();

    // Mouse is focused on input
    this._showFlyout = !!peopleInput.value;
  }

  /**
   * Queries the microsoft graph for a user based on the user id and adds them to the selectedPeople array
   *
   * @param {[string]} an array of user ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectUsersById(userIds: [string]): Promise<void> {
    const provider = Providers.globalProvider;
    const graph = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      // tslint:disable-next-line: forin
      for (const id in userIds) {
        try {
          const personDetails = await getUser(graph, userIds[id]);
          this.addPerson(personDetails);
          // tslint:disable-next-line: no-empty
        } catch (e) {}
      }
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update.
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  public render(): TemplateResult {
    const defaultTemplate = this.renderTemplate('default', { people: this.people });
    if (defaultTemplate) {
      return defaultTemplate;
    }

    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout();
    return html`
      <div class="people-picker">
        ${inputTemplate} ${flyoutTemplate}
      </div>
    `;
  }

  /**
   * Render the input text box.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderInput(): TemplateResult {
    const selectedPeopleTemplate = this.renderSelectedPeople();
    const inputTemplate = html`
      <div class="${this.selectedPeople.length ? 'input-search-start' : 'input-search'}">
        <input
          id="people-picker-input"
          class="people-selected-input"
          type="text"
          placeholder="Start typing a name"
          label="people-picker-input"
          aria-label="people-picker-input"
          role="input"
          .value="${this._userInput}"
          @keydown="${this.onUserKeyDown}"
          @keyup="${this.onUserKeyUp}"
        />
      </div>
    `;

    return html`
      <div class="people-picker-input" @click=${() => this.focus()}>
        <div class="people-selected-list">
          ${selectedPeopleTemplate} ${inputTemplate}
        </div>
      </div>
      <div class="people-list-separator"></div>
    `;
  }

  /**
   * Render the selected people tokens.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPeople(): TemplateResult {
    if (!this.selectedPeople || !this.selectedPeople.length) {
      return null;
    }

    return html`
      ${this.selectedPeople.slice(0, this.selectedPeople.length).map(
        person =>
          html`
            <div class="people-person">
              ${this.renderTemplate('selected-person', { person }, `selected-${person.id}`) ||
                this.renderSelectedPerson(person)}
              <div class="CloseIcon" @click="${() => this.removePerson(person)}">\uE711</div>
            </div>
          `
      )}
    `;
  }

  /**
   * Render the flyout chrome.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderFlyout(): TemplateResult {
    return html`
      <mgt-flyout .isOpen=${this._showFlyout}>
        <div slot="flyout" class="flyout">
          <div class="flyout-root">
            ${this.renderFlyoutContent()}
          </div>
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

    if (!this.people || this.people.length === 0 || this.showMax === 0) {
      return this.renderNoData();
    }

    const people = this.people.slice(0, this.showMax);
    (people[0] as any).isSelected = 'fill';

    return this.renderSearchResults(people);
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
          <div label="search-error-text" aria-label="loading" class="loading-text">
            ......
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
   * @memberof MgtPeoplePicker
   */
  protected renderNoData(): TemplateResult {
    return (
      this.renderTemplate('no-data', null, 'no-data') ||
      html`
        <div class="message-parent">
          <div label="search-error-text" aria-label="We didn't find any matches." class="search-error-text">
            We didn't find any matches.
          </div>
        </div>
      `
    );
  }

  /**
   * Render the list of search results.
   *
   * @protected
   * @param {any[]} people
   * @returns
   * @memberof MgtPeoplePicker
   */
  protected renderSearchResults(people: IDynamicPerson[]) {
    return html`
      <div class="people-list">
        ${repeat(
          people,
          person => person.id,
          person => html`
            <li
              class="list-person ${(person as any).isSelected === 'fill'
                ? 'people-person-list-fill'
                : 'people-person-list'}"
              @click="${() => this.onPersonClick(person)}"
            >
              ${this.renderPersonResult(person)}
            </li>
          `
        )}
      </div>
    `;
  }

  /**
   * Render an individual person search result.
   *
   * @protected
   * @param {MicrosoftGraph.Person} person
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderPersonResult(person: IDynamicPerson): TemplateResult {
    return (
      this.renderTemplate('person', { person }, person.id) ||
      html`
        <mgt-person .personDetails=${person} .personImage=${'@'}></mgt-person>
        <div class="people-person-text-area" id="${person.displayName}">
          ${this.renderHighlightText(person)}
          <span class="people-person-job-title">${person.jobTitle}</span>
        </div>
      `
    );
  }

  /**
   * Render an individual selected person token.
   *
   * @protected
   * @param {MicrosoftGraph.Person} person
   * @returns
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPerson(person: IDynamicPerson) {
    return html`
      <mgt-person
        class="selected-person"
        .personDetails=${person}
        .personImage=${'@'}
        show-name
        person-card="click"
      ></mgt-person>
    `;
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (this.groupId !== this._groupId) {
      this._groupId = this.groupId;
      const graph = provider.graph.forComponent(this);
      this.groupPeople = await getPeopleFromGroup(graph, this.groupId);
    }

    const query = this._userInput.toLowerCase();
    let people: any;

    // filtering groups
    if (this.groupId) {
      people = this.groupPeople;
    } else {
      const graph = provider.graph.forComponent(this);
      people = await findPerson(graph, query);
    }

    if (people) {
      people = people.filter((person: IDynamicPerson) => {
        return person.displayName.toLowerCase().indexOf(query) !== -1;
      });
    }

    this.people = this.filterPeople(people);
  }

  /**
   * Hide the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected hideFlyout(): void {
    this._showFlyout = false;
  }

  /**
   * Show the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected showFlyout(): void {
    this._showFlyout = true;
  }

  private renderHighlightText(person: IDynamicPerson) {
    let first: string = '';
    let last: string = '';
    let highlight: string = '';

    const displayName = person.displayName;
    const highlightLocation = displayName.toLowerCase().indexOf(this._userInput.toLowerCase());
    if (highlightLocation !== -1) {
      const userInputLength = this._userInput.length;

      // no location
      if (highlightLocation === 0) {
        // highlight is at the beginning of sentence
        first = '';
        highlight = displayName.slice(0, userInputLength);
        last = displayName.slice(userInputLength, displayName.length);
      } else if (highlightLocation === displayName.length) {
        // highlight is at end of the sentence
        first = displayName.slice(0, highlightLocation);
        highlight = displayName.slice(highlightLocation, displayName.length);
        last = '';
      } else {
        // highlight is in middle of sentence
        first = displayName.slice(0, highlightLocation);
        highlight = displayName.slice(highlightLocation, highlightLocation + userInputLength);
        last = displayName.slice(highlightLocation + userInputLength, displayName.length);
      }
    }

    return html`
      <div>
        <span class="people-person-text">${first}</span
        ><span class="people-person-text highlight-search-text">${highlight}</span
        ><span class="people-person-text">${last}</span>
      </div>
    `;
  }

  /**
   * Adds debounce method for set delay on user input
   */
  private onUserKeyUp(event: any) {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      return;
    }

    const input = event.target;

    if (event.code === 'Escape') {
      input.value = '';
      this._userInput = '';
      this.people = [];
      return;
    }
    if (event.code === 'Backspace' && this._userInput.length === 0 && this.selectedPeople.length > 0) {
      input.value = '';
      this._userInput = '';
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      // fire selected people changed event
      this.fireCustomEvent('selectionChanged', this.selectedPeople);
      return;
    }

    this.handleUserSearch(input);
  }

  private onPersonClick(person: IDynamicPerson) {
    this.addPerson(person);
    this.focus();
    this.hideFlyout();
  }

  /**
   * Tracks event on user input in search
   * @param input - input text
   */
  private handleUserSearch(input: any) {
    if (!this.debouncedSearch) {
      this.debouncedSearch = debounce(async () => {
        if (!this._userInput.length) {
          this.people = [];
          this.hideFlyout();
          this._showLoading = true;
        } else {
          // Wait a few milliseconds before showing the flyout.
          // This helps prevent loading state flickering while the user is actively changing the query.
          const loadingTimeout = setTimeout(() => {
            this._showLoading = true;
          }, 400);

          await this.loadState();
          clearTimeout(loadingTimeout);
          this._showLoading = false;
          this.showFlyout();
        }

        this.arrowSelectionCount = 0;
      }, 400);
    }

    if (this._userInput !== input.value) {
      this._userInput = input.value;
      this.debouncedSearch();
    }
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: any) {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this._userInput.length > 0) {
        event.preventDefault();
      }
    }
    if (event.code === 'Tab' || event.code === 'Enter') {
      if (this.people.length) {
        event.preventDefault();
      }
      this.addPerson(this.people[this.arrowSelectionCount]);
      this.hideFlyout();
      event.target.value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: any) {
    if (this.people.length) {
      // update arrow count
      if (event.keyCode === 38) {
        // up arrow
        if (this.arrowSelectionCount > 0) {
          this.arrowSelectionCount--;
        } else {
          this.arrowSelectionCount = 0;
        }
      }
      if (event.keyCode === 40) {
        // down arrow
        if (this.arrowSelectionCount + 1 !== this.people.length && this.arrowSelectionCount + 1 < this.showMax) {
          this.arrowSelectionCount++;
        } else {
          this.arrowSelectionCount = 0;
        }
      }

      const peopleList = this.renderRoot.querySelector('.people-list');
      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].setAttribute('class', 'list-person people-person-list');
      }
      // set selected background
      peopleList.children[this.arrowSelectionCount].setAttribute('class', 'list-person people-person-list-fill');
    }
  }

  /**
   * Tracks when user selects person from picker
   * @param person - contains details pertaining to selected user
   * @param event - tracks user event
   */
  private addPerson(person: IDynamicPerson) {
    if (person) {
      this._userInput = '';
      const duplicatePeople = this.selectedPeople.filter(p => {
        return p.id === person.id;
      });

      if (duplicatePeople.length === 0) {
        this.selectedPeople = [...this.selectedPeople, person];
        this.fireCustomEvent('selectionChanged', this.selectedPeople);

        this.people = [];
      }
    }
  }

  /**
   * Filters people searched from already selected people
   * @param people - array of people returned from query to Graph
   */
  private filterPeople(people: any) {
    // check if people need to be updated
    // ensuring people list is displayed
    // find ids from selected people
    if (people) {
      const idFilter = this.selectedPeople.map(el => {
        return el.id;
      });

      // filter id's
      const filtered = people.filter((person: IDynamicPerson) => {
        return idFilter.indexOf(person.id) === -1;
      });

      return filtered;
    }
  }

  /**
   * Removes person from selected people
   * @param person - person and details pertaining to user selected
   */
  private removePerson(person: IDynamicPerson) {
    const filteredPersonArr = this.selectedPeople.filter(p => {
      return p.id !== person.id;
    });
    this.selectedPeople = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this.selectedPeople);
  }

  private handleWindowClick(e: MouseEvent) {
    this.hideFlyout();
  }
}
