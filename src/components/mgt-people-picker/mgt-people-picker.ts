/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, property, query, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { findPerson, getPeopleFromGroup } from '../../graph/graph.people';
import { getUser } from '../../graph/graph.user';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { debounce } from '../../utils/Utils';
import { IDynamicPerson } from '../mgt-person/mgt-person';
import { PersonCardInteraction } from '../PersonCardInteraction';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-picker-css';

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
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({ attribute: 'group-id' })
  public get groupId(): string {
    return this._groupId;
  }
  public set groupId(value) {
    if (this._groupId === value) {
      return;
    }

    this._groupId = value;
    this.requestStateUpdate(true);
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
   *  array of user picked people.
   * @type {Array<IDynamicPerson>}
   */
  @property({
    attribute: 'selected-people',
    type: Array
  })
  public selectedPeople: IDynamicPerson[];

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  protected userInput: string;

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtLogin
   */
  @query('.flyout') protected flyout: MgtFlyout;

  // if search is still loading don't load "people not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  private _groupId: string;
  // tracking of user arrow key input for selection
  private _arrowSelectionCount: number = 0;
  // List of people requested if group property is provided
  private _groupPeople: IDynamicPerson[];
  private _debouncedSearch: { (): void; (): void };

  constructor() {
    super();

    this._showLoading = true;
    this._groupId = null;
    this.userInput = '';
    this.showMax = 6;
    this.selectedPeople = [];
  }

  /**
   * Invoked each time the custom element is appended into a document-connected element
   *
   * @memberof MgtLogin
   */
  public connectedCallback() {
    super.connectedCallback();
    this.addEventListener('click', e => e.stopPropagation());
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

    window.requestAnimationFrame(() => {
      // Mouse is focused on input
      if (!peopleInput.value) {
        this.hideFlyout();
      } else {
        this.showFlyout();
      }
    });
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

    const selectedPeopleTemplate = this.renderSelectedPeople(this.selectedPeople);
    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);
    return html`
      <div class="people-picker" @click=${() => this.focus()}>
        <div class="people-picker-input">
          <div class="people-selected-list">
            ${selectedPeopleTemplate} ${flyoutTemplate}
          </div>
        </div>
        <div class="people-list-separator"></div>
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
    const hasSelectedPeople = !!this.selectedPeople.length;
    const inputClasses = {
      'input-search': true,
      'input-search--start': hasSelectedPeople
    };

    return html`
      <div class="${classMap(inputClasses)}">
        <input
          id="people-picker-input"
          class="people-selected-input"
          type="text"
          placeholder="Start typing a name"
          label="people-picker-input"
          aria-label="people-picker-input"
          role="input"
          .value="${this.userInput}"
          @keydown="${this.onUserKeyDown}"
          @keyup="${this.onUserKeyUp}"
        />
      </div>
    `;
  }

  /**
   * Render the selected people tokens.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPeople(selectedPeople?: IDynamicPerson[]): TemplateResult {
    selectedPeople = selectedPeople || this.selectedPeople;
    if (!this.selectedPeople || !this.selectedPeople.length) {
      return null;
    }

    return html`
      ${selectedPeople.slice(0, selectedPeople.length).map(
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
  protected renderFlyout(anchor: TemplateResult): TemplateResult {
    return html`
      <mgt-flyout light-dismiss class="flyout">
        ${anchor}
        <div slot="flyout">
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
    (people[0] as IFocusable).isFocused = true;

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
      this.renderTemplate('error', null) ||
      this.renderTemplate('no-data', null) ||
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
   * @param {IDynamicPerson[]} people
   * @returns
   * @memberof MgtPeoplePicker
   */
  protected renderSearchResults(people?: IDynamicPerson[]) {
    people = people || this.people;

    return html`
      <div class="people-list">
        ${repeat(
          people,
          person => person.id,
          person => {
            const listPersonClasses = {
              focused: (person as IFocusable).isFocused,
              'list-person': true
            };
            return html`
              <li class="${classMap(listPersonClasses)}" @click="${() => this.onPersonClick(person)}">
                ${this.renderPersonResult(person)}
              </li>
            `;
          }
        )}
      </div>
    `;
  }

  /**
   * Render an individual person search result.
   *
   * @protected
   * @param {IDynamicPerson} person
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
   * @param {IDynamicPerson} person
   * @returns {TemplateResult}
   * @memberof MgtPeoplePicker
   */
  protected renderSelectedPerson(person: IDynamicPerson): TemplateResult {
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
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    const input = this.userInput.toLowerCase();
    let people: IDynamicPerson[];

    if (this.groupId) {
      const graph = provider.graph.forComponent(this);
      people = await getPeopleFromGroup(graph, this.groupId);
    } else if (input) {
      const graph = provider.graph.forComponent(this);
      people = await findPerson(graph, input);
    }

    if (people) {
      people = people.filter((person: IDynamicPerson) => {
        return person.displayName.toLowerCase().indexOf(input) !== -1;
      });
    }

    this.people = this.filterPeople(people);
  }

  /**
   * Show the flyout and its content.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
    }
  }

  /**
   * Dismiss the flyout.
   *
   * @protected
   * @memberof MgtLogin
   */
  protected hideFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.close();
    }
  }

  /**
   * Removes person from selected people
   * @param person - person and details pertaining to user selected
   */
  protected removePerson(person: IDynamicPerson): void {
    const filteredPersonArr = this.selectedPeople.filter(p => {
      return p.id !== person.id;
    });
    this.selectedPeople = filteredPersonArr;
    this.fireCustomEvent('selectionChanged', this.selectedPeople);
  }

  /**
   * Tracks when user selects person from picker
   * @param person - contains details pertaining to selected user
   * @param event - tracks user event
   */
  protected addPerson(person: IDynamicPerson): void {
    if (person) {
      this.userInput = '';
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

  private renderHighlightText(person: IDynamicPerson): TemplateResult {
    let first: string = '';
    let last: string = '';
    let highlight: string = '';

    const displayName = person.displayName;
    const highlightLocation = displayName.toLowerCase().indexOf(this.userInput.toLowerCase());
    if (highlightLocation !== -1) {
      const userInputLength = this.userInput.length;

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
  private onUserKeyUp(event: KeyboardEvent): void {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      return;
    }

    const input = event.target as HTMLInputElement;

    if (event.code === 'Escape') {
      input.value = '';
      this.userInput = '';
      this.people = [];
      return;
    }
    if (event.code === 'Backspace' && this.userInput.length === 0 && this.selectedPeople.length > 0) {
      input.value = '';
      this.userInput = '';
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      // fire selected people changed event
      this.fireCustomEvent('selectionChanged', this.selectedPeople);
      return;
    }

    this.handleUserSearch(input);
  }

  private onPersonClick(person: IDynamicPerson): void {
    this.addPerson(person);
    this.focus();
    this.hideFlyout();
  }

  /**
   * Tracks event on user input in search
   * @param input - input text
   */
  private handleUserSearch(input: HTMLInputElement) {
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        if (!this.userInput.length) {
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

        this._arrowSelectionCount = 0;
      }, 400);
    }

    if (this.userInput !== input.value) {
      this.userInput = input.value;
      this._debouncedSearch();
    }
  }

  /**
   * Tracks event on user search (keydown)
   * @param event - event tracked on user input (keydown)
   */
  private onUserKeyDown(event: KeyboardEvent): void {
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this.userInput.length > 0) {
        event.preventDefault();
      }
    }
    if (event.code === 'Tab' || event.code === 'Enter') {
      if (this.people.length) {
        event.preventDefault();
      }
      this.addPerson(this.people[this._arrowSelectionCount]);
      this.hideFlyout();
      (event.target as HTMLInputElement).value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: KeyboardEvent): void {
    if (this.people.length) {
      // update arrow count
      if (event.keyCode === 38) {
        // up arrow
        if (this._arrowSelectionCount > 0) {
          this._arrowSelectionCount--;
        } else {
          this._arrowSelectionCount = 0;
        }
      }
      if (event.keyCode === 40) {
        // down arrow
        if (this._arrowSelectionCount + 1 !== this.people.length && this._arrowSelectionCount + 1 < this.showMax) {
          this._arrowSelectionCount++;
        } else {
          this._arrowSelectionCount = 0;
        }
      }

      const peopleList = this.renderRoot.querySelector('.people-list');
      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].classList.remove('focused');
      }
      // set selected background
      peopleList.children[this._arrowSelectionCount].classList.add('focused');
    }
  }

  /**
   * Filters people searched from already selected people
   * @param people - array of people returned from query to Graph
   */
  private filterPeople(people: IDynamicPerson[]): IDynamicPerson[] {
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
}
