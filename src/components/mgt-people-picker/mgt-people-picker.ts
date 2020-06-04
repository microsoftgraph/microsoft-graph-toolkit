/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { User } from '@microsoft/microsoft-graph-types';
import { customElement, html, internalProperty, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { findGroups, GroupType } from '../../graph/graph.groups';
import { findPeople, getPeople, getPeopleFromGroup, PersonType } from '../../graph/graph.people';
import { findUsers, getUser, getUsersForUserIds } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import '../../styles/fabric-icon-font';
import { debounce } from '../../utils/Utils';
import { PersonViewType } from '../mgt-person/mgt-person';
import { PersonCardInteraction } from '../PersonCardInteraction';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-people-picker-css';

export { GroupType } from '../../graph/graph.groups';
export { PersonType } from '../../graph/graph.people';
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
 * @fires selectionChanged - Fired when selection changes
 *
 * @cssprop --font-color - {font} Default font color
 *
 * @cssprop --input-border - {String} Input section entire border
 * @cssprop --input-border-top - {String} Input section border top only
 * @cssprop --input-border-right - {String} Input section border right only
 * @cssprop --input-border-bottom - {String} Input section border bottom only
 * @cssprop --input-border-left - {String} Input section border left only
 * @cssprop --input-background-color - {Color} Input section background color
 * @cssprop --input-hover-color - {Color} Input text hover color
 * @cssprop --input-focus-color - {Color} Input text focus color
 *
 * @cssprop --dropdown-background-color - {Color} Background color of dropdown area
 * @cssprop --dropdown-item-hover-background - {Color} Background color of channel or team during hover
 * @cssprop --dropdown-item-selected-background - {Color} Background color of selected channel
 *
 * @cssprop --placeholder-focus-color - {Color} Color of placeholder text during focus state
 * @cssprop --placeholder-default-color - {Color} Color of placeholder text
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
   * value determining if search is filtered to a group.
   * @type {string}
   */
  @property({
    attribute: 'type',
    converter: (value, type) => {
      if (!value || value.length === 0) {
        return PersonType.Any;
      }

      if (typeof PersonType[value] === 'undefined') {
        return PersonType.Any;
      } else {
        return PersonType[value];
      }
    }
  })
  public get type(): PersonType {
    return this._type;
  }
  public set type(value) {
    if (this._type === value) {
      return;
    }

    this._type = value;
    this.requestStateUpdate(true);
  }

  /**
   * type of group to search for - requires personType to be
   * set to "Group" or "All"
   * @type {string}
   */
  @property({
    attribute: 'group-type',
    converter: (value, type) => {
      if (!value || value.length === 0) {
        return GroupType.Any;
      }

      const values = value.split(',');
      const groupTypes = [];

      for (let v of values) {
        v = v.trim();
        if (typeof GroupType[v] !== 'undefined') {
          groupTypes.push(GroupType[v]);
        }
      }

      if (groupTypes.length === 0) {
        return GroupType.Any;
      }

      // tslint:disable-next-line:no-bitwise
      return groupTypes.reduce((a, c) => a | c);
    }
  })
  public get groupType(): GroupType {
    return this._groupType;
  }
  public set groupType(value) {
    if (this._groupType === value) {
      return;
    }

    this._groupType = value;
    this.requestStateUpdate(true);
  }

  /**
   * containing object of IDynamicPerson.
   * @type {IDynamicPerson[]}
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
   * @type {IDynamicPerson[]}
   */
  @property({
    attribute: 'selected-people',
    type: Array
  })
  public selectedPeople: IDynamicPerson[];

  /**
   * array of people to be selected upon intialization
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'default-selected-user-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public defaultSelectedUserIds: string[];

  /**
   * Placeholder text.
   *
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'placeholder',
    type: String
  })
  public placeholder: string;

  /**
   * Determines whether component allows multiple or single selection of people
   *
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'selection-mode',
    type: String
  })
  public selectionMode: string;

  /**
   * User input in search.
   *
   * @protected
   * @type {string}
   * @memberof MgtPeoplePicker
   */
  protected userInput: string;

  // if search is still loading don't load "people not found" state
  @property({ attribute: false }) private _showLoading: boolean;

  private _groupId: string;
  private _type: PersonType = PersonType.Person;
  private _groupType: GroupType = GroupType.Any;

  private defaultPeople: IDynamicPerson[];

  // tracking of user arrow key input for selection
  private _arrowSelectionCount: number = 0;
  // List of people requested if group property is provided
  private _groupPeople: IDynamicPerson[];
  private _debouncedSearch: { (): void; (): void };
  @internalProperty() private _isFocused = false;

  @internalProperty() private _foundPeople: IDynamicPerson[];

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
    this.gainedFocus();
    const peopleInput = this.renderRoot.querySelector('.people-selected-input') as HTMLInputElement;
    if (!peopleInput) {
      return;
    }

    peopleInput.focus(options);
    peopleInput.select();

    // handles hiding control if default people have no more selections available
    const peopleLeft = this.filterPeople(this.defaultPeople);
    let shouldShow = true;
    if (peopleLeft && peopleLeft.length === 0) {
      shouldShow = false;
    }

    if (shouldShow) {
      window.requestAnimationFrame(() => {
        // Mouse is focused on input
        this.showFlyout();
      });
    }
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
    const defaultTemplate = this.renderTemplate('default', { people: this._foundPeople });
    if (defaultTemplate) {
      return defaultTemplate;
    }

    const selectedPeopleTemplate = this.renderSelectedPeople(this.selectedPeople);
    const inputTemplate = this.renderInput();
    const flyoutTemplate = this.renderFlyout(inputTemplate);

    const inputClasses = {
      focused: this._isFocused,
      'people-picker': true
    };
    return html`
      <div class=${classMap(inputClasses)} @click=${() => this.focus()}>
        <div class="people-picker-input">
          <div class="people-selected-list">
            ${selectedPeopleTemplate} ${flyoutTemplate}
          </div>
        </div>
      </div>
    `;
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
      this._groupPeople = null;
      this._foundPeople = null;
      this.selectedPeople = [];
    }

    return super.requestStateUpdate(force);
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

    const placeholder = this.placeholder ? this.placeholder : 'Start typing a name';

    const selectionMode = this.selectionMode ? this.selectionMode : 'multiple';

    const inputClasses = {
      'input-search': true,
      'input-search--start': hasSelectedPeople
    };

    if (selectionMode === 'single' && this.selectedPeople.length >= 1) {
      this.lostFocus();
      return null;
    }

    return html`
      <div class="${classMap(inputClasses)}">
        <input
          id="people-picker-input"
          class="people-selected-input"
          type="text"
          placeholder=${placeholder}
          label="people-picker-input"
          aria-label="people-picker-input"
          role="input"
          .value="${this.userInput}"
          @keydown="${this.onUserKeyDown}"
          @keyup="${this.onUserKeyUp}"
          @blur=${this.lostFocus}
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

    let people = this._foundPeople;

    if (!people || people.length === 0 || this.showMax === 0) {
      return this.renderNoData();
    }

    // clears focus
    for (const person of people) {
      (person as IFocusable).isFocused = false;
    }

    people = people.slice(0, this.showMax);
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
          <div class="spinner"></div>
          <div label="loading-text" aria-label="loading" class="loading-text">
            Loading...
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
    people = people || this._foundPeople;

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
    const user = person as User;
    const subTitle = user.jobTitle || user.mail;

    const classes = {
      'people-person-job-title': true,
      uppercase: !!user.jobTitle
    };

    return (
      this.renderTemplate('person', { person }, person.id) ||
      html`
        <mgt-person .personDetails=${person} .fetchImage=${true}></mgt-person>
        <div class="people-person-text-area" id="${person.displayName}">
          ${this.renderHighlightText(person)}
          <span class="${classMap(classes)}">${subTitle}</span>
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
        .fetchImage=${true}
        .view=${PersonViewType.oneline}
        .personCardInteraction=${PersonCardInteraction.click}
      ></mgt-person>
    `;
  }

  /**
   * Async query to Graph for members of group if determined by developer.
   * set's `this.groupPeople` to those members.
   */
  protected async loadState(): Promise<void> {
    let people = this.people;
    const input = this.userInput.toLowerCase();

    const provider = Providers.globalProvider;
    if (!people && provider && provider.state === ProviderState.SignedIn) {
      const graph = provider.graph.forComponent(this);

      // default common people
      if (!input.length && this._isFocused) {
        if (this.defaultPeople) {
          people = this.defaultPeople;
        } else {
          if (this.type === PersonType.Person || this.type === PersonType.Any) {
            people = await getPeople(graph);
            this.defaultPeople = people;
          }
        }
        this._showLoading = false;
      }

      if (this.defaultSelectedUserIds && !this.selectedPeople.length) {
        const defaultSelectedUsers = await getUsersForUserIds(graph, this.defaultSelectedUserIds);

        this.selectedPeople = [...defaultSelectedUsers];
        this.requestUpdate();
        this.fireCustomEvent('selectionChanged', this.selectedPeople);
      }

      if (this.groupId) {
        if (this._groupPeople === null) {
          try {
            this._groupPeople = await getPeopleFromGroup(graph, this.groupId);
          } catch (_) {
            this._groupPeople = [];
          }
        }

        people = this._groupPeople || [];
      } else if (!this.groupId && input) {
        people = [];
        if (this.type === PersonType.Person || this.type === PersonType.Any) {
          try {
            people = (await findPeople(graph, input, this.showMax)) || [];
          } catch (e) {
            // nop
          }

          if (people.length < this.showMax) {
            try {
              const users = (await findUsers(graph, input, this.showMax)) || [];

              // make sure only unique people
              const peopleIds = new Set(people.map(p => p.id));
              for (const user of users) {
                if (!peopleIds.has(user.id)) {
                  people.push(user);
                }
              }
            } catch (e) {
              // nop
            }
          }
        }
        // default group search with user input
        if ((this.type === PersonType.Group || this.type === PersonType.Any) && people.length < this.showMax) {
          people = [];
          try {
            const groups = (await findGroups(graph, '', this.showMax, this.groupType)) || [];
            people = people.concat(groups);
          } catch (e) {
            // nop
          }
        }
      }
      // default group search without user input
      if ((input.length === 0 && this.type === PersonType.Group) || this.type === PersonType.Any) {
        people = [];
        try {
          const groups = (await findGroups(graph, '', this.showMax, GroupType.Any)) || [];
          people = people.concat(groups);
        } catch (e) {
          // nop
        }
      }
    }

    if (people) {
      people = people.filter((user: User) => {
        return (
          user.displayName.toLowerCase().indexOf(input) !== -1 ||
          (!!user.givenName && user.givenName.toLowerCase().indexOf(input) !== -1) ||
          (!!user.surname && user.surname.toLowerCase().indexOf(input) !== -1) ||
          (!!user.mail && user.mail.toLowerCase().indexOf(input) !== -1)
        );
      });
    }

    this._foundPeople = this.filterPeople(people);
  }

  /**
   * Hide the results flyout.
   *
   * @protected
   * @memberof MgtPeoplePicker
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
   * @memberof MgtPeoplePicker
   */
  protected showFlyout(): void {
    const flyout = this.flyout;
    if (flyout) {
      flyout.open();
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
    this.loadState();
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
        if (!person.id) {
          return p.displayName === person.displayName;
        }
        return p.id === person.id;
      });

      if (duplicatePeople.length === 0) {
        this.selectedPeople = [...this.selectedPeople, person];
        this.fireCustomEvent('selectionChanged', this.selectedPeople);

        this.loadState();
        this._foundPeople = [];
      }
    }
  }

  private gainedFocus() {
    this._isFocused = true;
    const input = this.renderRoot.querySelector('.people-selected-input') as HTMLInputElement;
    if (input) {
      input.focus();
    }
    this._showLoading = true;
    this.loadState();
  }

  private lostFocus() {
    this._isFocused = false;
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
      this._foundPeople = [];
      return;
    }
    if (event.code === 'Backspace' && this.userInput.length === 0 && this.selectedPeople.length > 0) {
      input.value = '';
      this.userInput = '';
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      // reset flyout position
      this.hideFlyout();
      this.showFlyout();
      // fire selected people changed event
      this.fireCustomEvent('selectionChanged', this.selectedPeople);
      return;
    }

    this.handleUserSearch(input);
  }

  private onPersonClick(person: IDynamicPerson): void {
    this.addPerson(person);
    this.hideFlyout();
    if (this.selectionMode === 'single') {
      return;
    }
    this.focus();
  }

  /**
   * Tracks event on user input in search
   * @param input - input text
   */
  private handleUserSearch(input: HTMLInputElement) {
    this._showLoading = true;
    if (!this._debouncedSearch) {
      this._debouncedSearch = debounce(async () => {
        // Wait a few milliseconds before showing the flyout.
        // This helps prevent loading state flickering while the user is actively changing the query.

        const loadingTimeout = setTimeout(() => {
          if (!this.userInput.length) {
            this._foundPeople = [];
            this.hideFlyout();
          }
        }, 400);

        await this.loadState();
        clearTimeout(loadingTimeout);
        this._showLoading = false;
        this.showFlyout();

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
    if (!this.flyout.isOpen) {
      return;
    }
    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this.userInput.length > 0) {
        event.preventDefault();
      }
    }
    if (event.code === 'Tab' || event.code === 'Enter') {
      if (this._foundPeople.length) {
        event.preventDefault();
      }
      this.addPerson(this._foundPeople[this._arrowSelectionCount]);
      this.hideFlyout();
      (event.target as HTMLInputElement).value = '';
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event: KeyboardEvent): void {
    if (this._foundPeople.length) {
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
        if (
          this._arrowSelectionCount + 1 !== this._foundPeople.length &&
          this._arrowSelectionCount + 1 < this.showMax
        ) {
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
        return el.id ? el.id : el.displayName;
      });

      // filter id's
      const filtered = people.filter((person: IDynamicPerson) => {
        if (person.id) {
          return idFilter.indexOf(person.id) === -1;
        } else {
          return idFilter.indexOf(person.displayName) === -1;
        }
      });

      return filtered;
    }
  }
}
