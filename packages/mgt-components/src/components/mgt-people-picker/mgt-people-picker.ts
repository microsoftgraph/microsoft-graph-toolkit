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
import {
  findGroups,
  getGroupsForGroupIds,
  GroupType,
  getGroup,
  findGroupsFromGroupIds
} from '../../graph/graph.groups';
import { findPeople, getPeople, PersonType, UserType } from '../../graph/graph.people';
import {
  findUsers,
  findGroupMembers,
  findUsersFromGroupIds,
  getUser,
  getUsersForUserIds,
  getUsers
} from '../../graph/graph.user';
import { IDynamicPerson, ViewType } from '../../graph/types';
import { Providers, ProviderState, MgtTemplatedComponent, arraysAreEqual, IGraph } from '@microsoft/mgt-element';
import '../../styles/style-helper';
import '../sub-components/mgt-spinner/mgt-spinner';
import { debounce, isValidEmail } from '../../utils/Utils';
import { MgtPerson } from '../mgt-person/mgt-person';
import { PersonCardInteraction } from '../PersonCardInteraction';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { styles } from './mgt-people-picker-css';

import { strings } from './strings';

export { GroupType } from '../../graph/graph.groups';
export { PersonType, UserType } from '../../graph/graph.people';

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
 * @cssprop --color - {Color} Default font color
 *
 * @cssprop --input-border - {String} Input section entire border
 * @cssprop --input-border-top - {String} Input section border top only
 * @cssprop --input-border-right - {String} Input section border right only
 * @cssprop --input-border-bottom - {String} Input section border bottom only
 * @cssprop --input-border-left - {String} Input section border left only
 * @cssprop --input-background-color - {Color} Input section background color
 * @cssprop --input-border-color--hover - {Color} Input border hover color
 * @cssprop --input-border-color--focus - {Color} Input border focus color
 *
 * @cssprop --selected-person-background-color - {Color} Selected person background color
 *
 * @cssprop --dropdown-background-color - {Color} Background color of dropdown area
 * @cssprop --dropdown-item-hover-background - {Color} Background color of person during hover
 *
 * @cssprop --placeholder-color--focus - {Color} Color of placeholder text during focus state
 * @cssprop --placeholder-color - {Color} Color of placeholder text
 *
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
  @property({ attribute: 'group-id', converter: value => value.trim() })
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
   * value determining if search is filtered to multiple groups.
   * @type {string}
   */
  @property({
    attribute: 'group-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get groupIds(): string[] {
    return this._groupIds;
  }
  public set groupIds(value) {
    if (arraysAreEqual(this._groupIds, value)) {
      return;
    }
    this._groupIds = value;
    this.requestStateUpdate(true);
  }

  /**
   * value determining if search is filtered to a group.
   * @type {PersonType}
   */
  @property({
    attribute: 'type',
    converter: (value, type) => {
      value = value.toLowerCase();
      if (!value || value.length === 0) {
        return PersonType.any;
      }

      if (typeof PersonType[value] === 'undefined') {
        return PersonType.any;
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
   * @type {GroupType}
   */
  @property({
    attribute: 'group-type',
    converter: (value, type) => {
      if (!value || value.length === 0) {
        return GroupType.any;
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
        return GroupType.any;
      }

      // tslint:disable-next-line:no-bitwise
      const gt = groupTypes.reduce((a, c) => a | c);
      return gt;
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

  @property({
    attribute: 'user-type',
    converter: (value, type) => {
      value = value.toLowerCase();
      if (!value || value.length === 0) {
        return UserType.any;
      }

      if (typeof UserType[value] === 'undefined') {
        return UserType.any;
      } else {
        return UserType[value];
      }
    }
  })
  public get userType(): UserType {
    return this._userType;
  }
  public set userType(value) {
    if (this._userType === value) {
      return;
    }

    this._userType = value;
    this.requestStateUpdate(true);
  }

  /**
   * whether the return should contain a flat list of all nested members
   * @type {boolean}
   */
  @property({
    attribute: 'transitive-search',
    type: Boolean
  })
  public transitiveSearch: boolean;

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
   * Sets whether the person image should be fetched
   * from the Microsoft Graph
   *
   * @type {boolean}
   * @memberof MgtPerson
   */
  @property({
    attribute: 'disable-images',
    type: Boolean
  })
  public disableImages: boolean;

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
   * array of groups to be selected upon intialization
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'default-selected-group-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public defaultSelectedGroupIds: string[];

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

  /**
   * Determines if a user can enter an email without selecting a person
   *
   * @type {boolean}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'allow-any-email',
    type: Boolean
  })
  public allowAnyEmail: boolean;

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

  private _userIds: string[];
  /**
   * Array of the only users to be searched.
   *
   * @type {string[]}
   * @memberof MgtPeoplePicker
   */
  @property({
    attribute: 'user-ids',
    converter: value => {
      return value.split(',').map(v => v.trim());
    },
    type: String
  })
  public get userIds(): string[] {
    return this._userIds;
  }
  public set userIds(value: string[]) {
    if (arraysAreEqual(this._userIds, value)) {
      return;
    }
    this._userIds = value;
    this.requestStateUpdate(true);
  }

  /**
   * Filters that can be set on the user properties query.
   */
  @property({ attribute: 'user-filters' })
  public get userFilters(): string {
    return this._userFilters;
  }

  public set userFilters(value: string) {
    this._userFilters = value;
    this.requestStateUpdate(true);
  }

  /**
   * Filters that can be set on the people query properties.
   */
  @property({ attribute: 'people-filters' })
  public get peopleFilters(): string {
    return this._peopleFilters;
  }

  public set peopleFilters(value: string) {
    this._peopleFilters = value;
    this.requestStateUpdate(true);
  }

  /**
   * Filters that can be set on the group query properties.
   */
  @property({ attribute: 'group-filters' })
  public get groupFilters(): string {
    return this._groupFilters;
  }

  public set groupFilters(value: string) {
    this._groupFilters = value;
    this.requestStateUpdate(true);
  }

  /**
   * Get the scopes required for people picker
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtPeoplePicker
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set(['user.read.all', 'people.read', 'group.read.all', 'user.readbasic.all', ...MgtPerson.requiredScopes])
    ];
  }

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
  private _groupIds: string[];
  private _type: PersonType = PersonType.person;
  private _groupType: GroupType = GroupType.any;
  private _userType: UserType = UserType.any;
  private _currentSelectedUser: IDynamicPerson;
  private _userFilters: string;
  private _groupFilters: string;
  private _peopleFilters: string;

  private defaultPeople: IDynamicPerson[];

  // tracking of user arrow key input for selection
  private _arrowSelectionCount: number = 0;
  // List of people requested if group property is provided
  private _groupPeople: IDynamicPerson[];
  private _debouncedSearch: { (): void; (): void };
  private defaultSelectedUsers: IDynamicPerson[];
  private defaultSelectedGroups: IDynamicPerson[];
  // List of users highlighted for copy/cut-pasting
  private _highlightedUsers: Element[] = [];
  // current user index to the left of the highlighted users
  private _currentHighlightedUserPos: number = 0;

  @internalProperty() private _isFocused = false;

  @internalProperty() private _foundPeople: IDynamicPerson[];

  private _mouseLeaveTimeout;
  private _mouseEnterTimeout;
  private _isKeyboardFocus: boolean = true;

  constructor() {
    super();
    this.clearState();
    this._showLoading = true;
    this.showMax = 6;
    this.disableImages = false;

    this.disabled = false;
    this.allowAnyEmail = false;
    this.addEventListener('copy', this.handleCopy);
    this.addEventListener('cut', this.handleCut);
    this.addEventListener('paste', this.handlePaste);
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
   * Queries the microsoft graph for a user based on the user id and adds them to the selectedPeople array
   *
   * @param {readonly string []} an array of user ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectUsersById(userIds: readonly string[]): Promise<void> {
    const provider = Providers.globalProvider;
    const graph = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      // tslint:disable-next-line: forin
      for (const id in userIds) {
        const userId = userIds[id];
        try {
          const personDetails = await getUser(graph, userId);
          this.addPerson(personDetails);
        } catch (e) {
          // This caters for allow-any-email property if it's enabled on the component
          if (e.message && e.message.includes('does not exist') && this.allowAnyEmail) {
            if (isValidEmail(userId)) {
              const anyMailUser = {
                mail: userId,
                displayName: userId
              };
              this.addPerson(anyMailUser);
            }
          }
        }
      }
    }
  }
  /**
   * Queries the microsoft graph for a group of users from a group id, and adds them to the selectedPeople
   *
   * @param {readonly string []} an array of group ids to add to selectedPeople
   * @returns {Promise<void>}
   * @memberof MgtPeoplePicker
   */
  public async selectGroupsById(groupIds: readonly string[]): Promise<void> {
    const provider = Providers.globalProvider;
    const graph = Providers.globalProvider.graph;
    if (provider && provider.state === ProviderState.SignedIn) {
      // tslint:disable-next-line: forin
      for (const id in groupIds) {
        try {
          const groupDetails = await getGroup(graph, groupIds[id]);
          this.addPerson(groupDetails);
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
      'people-picker': true,
      disabled: this.disabled
    };

    return html`
       <div dir=${this.direction} class=${classMap(inputClasses)} @click=${e => this.focus(e)}>
         <div
          aria-expanded="false"
          aria-haspopup="listbox"
          role="combobox"
          class="selected-list"
          id="selected-list">
           ${selectedPeopleTemplate} ${flyoutTemplate}
         </div>
       </div>
     `;
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtPeoplePicker
   */
  protected clearState(): void {
    // this._groupId = null;
    this.selectedPeople = [];
    this.userInput = '';
    this._highlightedUsers = [];
    this._currentHighlightedUserPos = 0;
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

    const placeholder = !this.disabled
      ? this.placeholder
        ? this.placeholder
        : this.strings.inputPlaceholderText
      : this.placeholder || '';

    const selectionMode = this.selectionMode ? this.selectionMode : 'multiple';

    const inputClasses = {
      'search-box': true,
      'search-box-start': hasSelectedPeople
    };

    if (selectionMode === 'single' && this.selectedPeople.length >= 1) {
      this.lostFocus();
      return null;
    }

    const inputAriaLabelText = `${
      this._currentSelectedUser !== undefined
        ? this.strings.selected + ' ' + this._currentSelectedUser.displayName + ' '
        : ''
    } ' people-picker-input'`;

    return html`
       <div class="${classMap(inputClasses)}">
         <input
           id="people-picker-input"
           class="search-box__input"
           type="text"
           role="combobox"
           placeholder=${placeholder}
           label="people-picker-input"
           autocomplete="off"
           aria-label=${inputAriaLabelText}
           aria-autocomplete="list"
           aria-expanded="false"
           tabindex="0"
           @keydown="${this.onUserKeyDown}"
           @keyup="${this.onUserKeyUp}"
           @blur=${this.lostFocus}
           @click=${this.handleFlyout}
           ?disabled=${this.disabled}
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
       <div
        tabindex="0"
        aria-label="selected-people"
        aria-orientation="vertical"
        role="listbox"
        class="selected-list__options">${selectedPeople.slice(0, selectedPeople.length).map(
          person =>
            html`
             <div
             role="option"
             tabindex="0"
             aria-label=${person.displayName}
             class="selected-list__person-wrapper">
               ${
                 this.renderTemplate(
                   'selected-person',
                   { person },
                   `selected-${person.id ? person.id : person.displayName}`
                 ) || this.renderSelectedPerson(person)
               }

               <div class="selected-list__person-wrapper__overflow">
                 <div class="selected-list__person-wrapper__overflow__gradient"></div>
                 <div
                   tabindex="0"
                   aria-label="close-icon"
                   class="selected-list__person-wrapper__overflow__close-icon"
                   @click="${e => this.removePerson(person, e)}"
                   @keydown="${e => this.removePerson(person, e)}"
                 >
                   \uE711
                 </div>
               </div>
             </div>
           `
        )}</div>
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
   * @memberof MgtPeoplePicker
   */
  protected renderNoData(): TemplateResult {
    if (!this._isFocused) {
      return;
    }
    return (
      this.renderTemplate('error', null) ||
      this.renderTemplate('no-data', null) ||
      html`
         <div class="message-parent">
           <div label="search-error-text" aria-label=${this.strings.noResultsFound} class="search-error-text">
             ${this.strings.noResultsFound}
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
    let filteredPeople = people.filter(person => person.id);
    let firstName = '';

    const selectedList = this.renderRoot.querySelector('.selected-list');
    let names = '';
    for (let i = 0; i < filteredPeople.length; i++) {
      names += filteredPeople[i].displayName.toString() + ' ';
    }

    return html`
      <div
        id="suggestions-list"
        class="people-list"
        aria-expanded="true"
        role="list"
        aria-label="people-picker-input input text ${
          this.userInput.length === 0 ? this.strings.inputPlaceholderText : this.userInput
        } ${this.strings.suggestedContacts} ${names}"
        @mouseenter=${this.handleMouseEnter}
        @mouseleave=${this.handleMouseLeave}>
         ${repeat(
           filteredPeople,
           person => person.id,
           person => {
             const listPersonClasses = {
               focused:
                 this._isKeyboardFocus && this._arrowSelectionCount === 0 ? (person as IFocusable).isFocused : '',
               'list-person': true
             };
             return html`
               <li
                role="option"
                @keydown="${this.onUserKeyDown}"
                aria-label=" ${this.strings.suggestedContact} ${person.displayName}"
                id="${person.displayName}"
                tabindex="0"
                class="${classMap(listPersonClasses)}"
                @click="${e => this.onPersonClick(person)}">
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
         <mgt-person .personDetails=${person} .fetchImage=${!this.disableImages}></mgt-person>
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
         tabindex="-1"
         class="selected-list__person-wrapper__person"
         .personDetails=${person}
         .fetchImage=${!this.disableImages}
         .view=${ViewType.oneline}
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

      if (!input.length && this._isFocused) {
        if (this.defaultPeople) {
          people = this.defaultPeople;
        } else {
          if (this.groupId || this.groupIds) {
            if (this._groupPeople === null) {
              if (this.groupId) {
                try {
                  if (this.type === PersonType.group) {
                    this._groupPeople = await findGroupMembers(
                      graph,
                      null,
                      this.groupId,
                      this.showMax,
                      this.type,
                      this.transitiveSearch
                    );
                  } else {
                    this._groupPeople = await findGroupMembers(
                      graph,
                      null,
                      this.groupId,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters,
                      this.peopleFilters
                    );
                  }
                } catch (_) {
                  this._groupPeople = [];
                }
              } else if (this.groupIds) {
                if (this.type === PersonType.group) {
                  try {
                    this._groupPeople = await getGroupsForGroupIds(graph, this.groupIds, this.groupFilters);
                  } catch (_) {
                    this._groupPeople = [];
                  }
                } else {
                  try {
                    const peopleInGroups = await findUsersFromGroupIds(
                      graph,
                      '',
                      this.groupIds,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters
                    );
                    this._groupPeople = peopleInGroups;
                  } catch (_) {
                    this._groupPeople = [];
                  }
                }
              }
            }
            people = this._groupPeople || [];
          } else if (this.type === PersonType.person || this.type === PersonType.any) {
            if (this.userIds) {
              people = await getUsersForUserIds(graph, this.userIds, '', this.userFilters);
            } else {
              const isUserOrContactType = this.userType === UserType.user || this.userType === UserType.contact;
              if (this._userFilters && isUserOrContactType) {
              } else {
                people = await getPeople(graph, this.userType, this._peopleFilters);
              }
            }
            people = await getUsers(graph, this._userFilters, this.showMax);
          } else if (this.type === PersonType.group) {
            if (this.groupIds) {
              try {
                people = await this.getGroupsForGroupIds(graph, people);
              } catch (_) {}
            } else {
              let groups = (await findGroups(graph, '', this.showMax, this.groupType, this._groupFilters)) || [];
              if (groups.length > 0 && groups[0]['value']) {
                groups = groups[0]['value'];
              }
              people = groups;
            }
          }
          this.defaultPeople = people;
        }
      }
      this._showLoading = false;

      if (
        (this.defaultSelectedUserIds || this.defaultSelectedGroupIds) &&
        !this.selectedPeople.length &&
        !this.defaultSelectedUsers
      ) {
        this.defaultSelectedUsers = await getUsersForUserIds(graph, this.defaultSelectedUserIds, '', this.userFilters);
        this.defaultSelectedGroups = await getGroupsForGroupIds(
          graph,
          this.defaultSelectedGroupIds,
          this.peopleFilters
        );

        this.defaultSelectedGroups = this.defaultSelectedGroups.filter(group => {
          return group !== null;
        });

        this.defaultSelectedUsers = this.defaultSelectedUsers.filter(user => {
          return user !== null;
        });

        this.selectedPeople = [...this.defaultSelectedUsers, ...this.defaultSelectedGroups];
        this.requestUpdate();
        this.fireCustomEvent('selectionChanged', this.selectedPeople);
      }

      if (input) {
        people = [];

        if (this.groupId) {
          people =
            (await findGroupMembers(
              graph,
              input,
              this.groupId,
              this.showMax,
              this.type,
              this.transitiveSearch,
              this.userFilters,
              this.peopleFilters
            )) || [];
        } else {
          if (this.type === PersonType.person || this.type === PersonType.any) {
            try {
              // Default UserType === any
              if (this.userType === UserType.contact || this.userType === UserType.user) {
                // we might have a user-filters property set, search for users with it.
                if (this.userIds && this.userIds.length) {
                  // has the user-ids proerty set
                  people = await getUsersForUserIds(graph, this.userIds, input, this._userFilters);
                } else {
                  people = await findUsers(graph, input, this.showMax, this._userFilters);
                }
              } else {
                if (!this.groupIds) {
                  if (this.userIds && this.userIds.length) {
                    // has the user-ids proerty set
                    people = await getUsersForUserIds(graph, this.userIds, input, this._userFilters);
                  } else {
                    people = (await findPeople(graph, input, this.showMax, this.userType, this._peopleFilters)) || [];
                  }
                } else {
                  // Does not work when the PersonType = person.
                  try {
                    people = await findUsersFromGroupIds(
                      graph,
                      input,
                      this.groupIds,
                      this.showMax,
                      this.type,
                      this.transitiveSearch,
                      this.userFilters
                    );
                  } catch (_) {}
                }
              }
            } catch (e) {
              // nop
            }

            // Don't follow this path if a people-filters attribute is set on the component as the
            // default type === PersonType.person
            if (
              people &&
              people.length < this.showMax &&
              this.userType !== UserType.contact &&
              this.type !== PersonType.person
            ) {
              try {
                const users = (await findUsers(graph, input, this.showMax, this._userFilters)) || [];

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

          if ((this.type === PersonType.group || this.type === PersonType.any) && people.length < this.showMax) {
            if (this.groupIds) {
              try {
                people = await findGroupsFromGroupIds(
                  graph,
                  input,
                  this.groupIds,
                  this.showMax,
                  this.groupType,
                  this.userFilters
                );
              } catch (_) {}
            } else {
              let groups = [];
              try {
                groups = (await findGroups(graph, input, this.showMax, this.groupType, this._groupFilters)) || [];
                people = people.concat(groups);
              } catch (e) {
                // nop
              }
            }
          }
        }
      }
    }
    //people = this.getUniquePeople(people);
    this._foundPeople = this.filterPeople(people);
  }

  /**
   * Gets the Groups in a list of group IDs.
   *
   * @param graph the graph object
   * @param people already found groups
   * @returns groups found
   */
  private async getGroupsForGroupIds(graph: IGraph, people: IDynamicPerson[]) {
    const groups = await getGroupsForGroupIds(graph, this.groupIds, this.groupFilters);
    for (let group of groups as IDynamicPerson[]) {
      people = people.concat(group);
    }
    people = people.filter(person => person);
    return people;
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
  protected removePerson(person: IDynamicPerson, e: MouseEvent): void {
    e.stopPropagation();
    const filteredPersonArr = this.selectedPeople.filter(p => {
      if (!person.id && p.displayName) {
        return p.displayName !== person.displayName;
      }
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
      setTimeout(() => {
        this.clearInput();
      }, 50);
      const duplicatePeople = this.selectedPeople.filter(p => {
        if (!person.id && p.displayName) {
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

  private clearInput() {
    this.clearHighlighted();
    if (this.selectionMode !== 'single') {
      this.input.value = '';
    }
    this.userInput = '';
  }

  private handleFlyout() {
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

  private gainedFocus() {
    this.clearHighlighted();
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

  private handleMouseEnter(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseEnterTimeout = setTimeout(this.hideKeyboardFocus.bind(this), 100);
  }

  private handleMouseLeave(e: MouseEvent) {
    clearTimeout(this._mouseEnterTimeout);
    clearTimeout(this._mouseLeaveTimeout);
    this._mouseLeaveTimeout = setTimeout(this.showKeyboardFocus.bind(this), 100);
  }

  private hideKeyboardFocus() {
    this._isKeyboardFocus = false;
    this.requestUpdate();

    const peopleList = this.renderRoot.querySelector('.people-list');
    if (peopleList && peopleList.children.length) {
      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].classList.remove('focused');
      }
    }
  }

  private showKeyboardFocus() {
    this._isKeyboardFocus = true;
    this.handleArrowSelection();
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
    } else {
      first = person.displayName;
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
    const isCmdOrCtrlKey = ['ControlLeft', 'ControlRight'].includes(event.code) || event.ctrlKey || event.metaKey;
    const isArrowKey = ['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft'].includes(event.code);

    if (isCmdOrCtrlKey || isArrowKey) {
      if (isCmdOrCtrlKey || ['ArrowLeft', 'ArrowRight'].includes(event.code)) {
        // Only hide the flyout when you're doing selections with Left/Right Arrow key
        this.hideFlyout();
      }
      return;
    }

    if (event.code === 'Tab' && !this.flyout.isOpen) {
      // keyCodes capture: tab (9)
      if (this.allowAnyEmail) {
        this.gainedFocus();
      }
    }

    if (event.shiftKey) {
      this.gainedFocus();
    }

    const input = event.target as HTMLInputElement;

    if (event.code === 'Escape') {
      this.clearInput();
      this._foundPeople = [];
    } else if (event.code === 'Backspace' && this.userInput.length === 0 && this.selectedPeople.length > 0) {
      this.clearHighlighted();
      // remove last person in selected list
      this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - 1);
      this.loadState();
      this.hideFlyout();
      // fire selected people changed event
      this.fireCustomEvent('selectionChanged', this.selectedPeople);
    } else if (event.code === 'Comma' || event.code === 'Semicolon') {
      if (this.allowAnyEmail) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    } else {
      this.userInput = input.value;
      const validEmail = isValidEmail(this.userInput);
      if (!validEmail) {
        this.handleUserSearch();
      }
    }
  }

  private handleAnyEmail() {
    this._showLoading = false;
    this._arrowSelectionCount = 0;
    if (isValidEmail(this.userInput)) {
      const anyMailUser = {
        mail: this.userInput,
        displayName: this.userInput
      };
      this.addPerson(anyMailUser);
    }
    this.hideFlyout();
    if (this.input) {
      this.input.focus();
      this._isFocused = true;
    }
  }

  private onPersonClick(person: IDynamicPerson): void {
    this._currentSelectedUser = person;
    this.addPerson(person);
    this.hideFlyout();

    if (!this.input) {
      return;
    }

    this.input.focus();
    this._isFocused = true;
    this.hideFlyout();
    if (this.selectionMode === 'single') {
      return;
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
    const selectedList = this.renderRoot.querySelector('.selected-list');
    const isCmdOrCtrlKey = event.ctrlKey || event.metaKey;
    if (isCmdOrCtrlKey && selectedList) {
      const selectedPeople = selectedList.querySelectorAll('mgt-person.selected-list__person-wrapper__person');
      this.hideFlyout();
      if (isCmdOrCtrlKey && event.code === 'ArrowLeft') {
        this._currentHighlightedUserPos =
          (this._currentHighlightedUserPos - 1 + selectedPeople.length) % selectedPeople.length;
        if (this._currentHighlightedUserPos >= 0 && this._currentHighlightedUserPos !== NaN) {
          this._highlightedUsers.push(selectedPeople[this._currentHighlightedUserPos]);
        } else {
          this._currentHighlightedUserPos = 0;
        }
      } else if (isCmdOrCtrlKey && event.code === 'ArrowRight') {
        const person = this._highlightedUsers.pop();
        if (person) {
          const personParent = person.parentElement;
          if (personParent) {
            this.clearHighlighted(personParent);
            this._currentHighlightedUserPos++;
          }
        }
      } else if (isCmdOrCtrlKey && event.code === 'KeyA') {
        this._highlightedUsers = [];
        selectedPeople.forEach(person => this._highlightedUsers.push(person));
      }
      if (this._highlightedUsers) {
        this.highlightSelectedPeople(this._highlightedUsers);
      }
      return;
    }

    this.clearHighlighted();

    if (!this.flyout.isOpen) {
      return;
    }

    if (event.keyCode === 40 || event.keyCode === 38) {
      // keyCodes capture: down arrow (40) and up arrow (38)
      this.handleArrowSelection(event);
      if (this.input.value.length > 0) {
        event.preventDefault();
      }
    }

    const input = event.target as HTMLInputElement;
    if (event.code === 'Tab' || event.code === 'Enter') {
      if (!event.shiftKey && this._foundPeople) {
        // keyCodes capture: tab (9) and enter (13)
        event.preventDefault();
        event.stopPropagation();
        if (this._foundPeople.length) {
          this.fireCustomEvent('blur');
        }

        const foundPerson = this._foundPeople[this._arrowSelectionCount];
        if (foundPerson) {
          this.addPerson(foundPerson);
        } else if (this.allowAnyEmail) {
          this.handleAnyEmail();
        }
      }
      this.hideFlyout();
      (event.target as HTMLInputElement).value = '';
    } else if (event.code === 'Comma' || event.code === 'Semicolon') {
      if (this.allowAnyEmail) {
        event.preventDefault();
        event.stopPropagation();
        this.userInput = input.value;
        this.handleAnyEmail();
      }
    }
  }

  /**
   * Gets the text of the highlighed people and writes it to the clipboard
   */
  private async writeHighlightedText() {
    const copyText = [];
    for (let i = 0; i < this._highlightedUsers.length; i++) {
      const element: any = this._highlightedUsers[i];
      const _personDetails = element._personDetails;
      const { id, displayName, email, userPrincipalName, scoredEmailAddresses } = _personDetails;
      let emailAddress: string;
      if (scoredEmailAddresses && scoredEmailAddresses.length > 0) {
        emailAddress = scoredEmailAddresses.pop().address;
      } else {
        emailAddress = userPrincipalName || email;
      }

      copyText.push({ id, displayName, email: emailAddress });
    }
    let copiedTextStr: string = '';
    if (copyText.length > 0) {
      copiedTextStr = JSON.stringify(copyText);
    }

    await navigator.clipboard.writeText(copiedTextStr);
  }

  /**
   * Handles the cut event when it is fired
   */
  private async handleCut() {
    await this.writeHighlightedText();
    this.removeHighlightedOnCut();
  }

  /**
   * Handles the copy event when it is fired
   */
  private async handleCopy() {
    await this.writeHighlightedText();
  }

  /**
   * Parses the copied people text and adds them when you paste
   */
  private async handlePaste() {
    try {
      const copiedText = await navigator.clipboard.readText();
      if (copiedText) {
        try {
          const people = JSON.parse(copiedText);
          if (people && people.length > 0) {
            for (const person of people) {
              this.addPerson(person);
            }
          }
        } catch (error) {
          if (error instanceof SyntaxError) {
            const _delimeters = [',', ';'];
            let listOfUsers: Array<string>;
            try {
              for (let i = 0; i < _delimeters.length; i++) {
                listOfUsers = copiedText.split(_delimeters[i]);
                if (listOfUsers.length > 1) {
                  this.hideFlyout();
                  this.selectUsersById(listOfUsers);
                  break;
                }
              }
              // tslint:disable-next-line: no-empty
            } catch (error) {}
          }
        }
      }
    } catch (error) {
      // 'navigator.clipboard.readText is not a function' error is thrown in Mozilla
      // more information here https://developer.mozilla.org/en-US/docs/Web/API/Clipboard/readText#browser_compatibility
      // Firefox only supports reading the clipboard in browser extensions,
      // using the "clipboardRead" extension permission.
    }
  }

  /**
   * Removes only the highlighted elements from the peoplePicker during cut operations.
   */
  private removeHighlightedOnCut() {
    this.selectedPeople = this.selectedPeople.splice(0, this.selectedPeople.length - this._highlightedUsers.length);
    this._highlightedUsers = [];
    this._currentHighlightedUserPos = 0;
    this.loadState();
    this.hideFlyout();
    this.fireCustomEvent('selectionChanged', this.selectedPeople);
  }
  /**
   * Changes the color class to show which people are selected for copy/cut-paste
   * @param people list of selected people classes
   */
  private highlightSelectedPeople(people: Element[]) {
    for (let i = 0; i < people.length; i++) {
      const person = people[i];
      const parentElement = person.parentElement;
      parentElement.setAttribute('class', 'selected-list__person-wrapper-highlighted');

      const personNodes = Array.from(parentElement.getElementsByClassName('selected-list__person-wrapper__person'));
      if (personNodes && personNodes.length > 0) {
        const personNode = personNodes.pop();
        personNode.setAttribute('class', 'selected-list__person-wrapper-highlighted__person');
      }

      const gradientNodes = Array.from(
        parentElement.getElementsByClassName('selected-list__person-wrapper__overflow__gradient')
      );
      if (gradientNodes && gradientNodes.length > 0) {
        const gradientNode = gradientNodes.pop();
        gradientNode.setAttribute('class', 'selected-list__person-wrapper-highlighted__overflow__gradient');
      }

      const closeIconNodes = Array.from(
        parentElement.getElementsByClassName('selected-list__person-wrapper__overflow__close-icon')
      );
      if (closeIconNodes && closeIconNodes.length > 0) {
        const closeIconNode = closeIconNodes.pop();
        closeIconNode.setAttribute('class', 'selected-list__person-wrapper-highlighted__overflow__close-icon');
      }
    }
  }

  /**
   * Defaults the people class back to the normal view
   */
  private clearHighlighted(node?: Element) {
    if (node) {
      this.clearNodeHighlights(node);
    } else {
      for (let i = 0; i < this._highlightedUsers.length; i++) {
        const person = this._highlightedUsers[i];
        const parentElement = person.parentElement;
        if (parentElement) {
          this.clearNodeHighlights(parentElement);
        }
      }
      this._highlightedUsers = [];
      this._currentHighlightedUserPos = 0;
    }
  }

  /**
   * Returns the original classes of a highlighted person element
   * @param node a highlighted node element
   */
  private clearNodeHighlights(node: Element) {
    node.setAttribute('class', 'selected-list__person-wrapper');

    const personNodes = Array.from(node.getElementsByClassName('selected-list__person-wrapper-highlighted__person'));
    if (personNodes && personNodes.length > 0) {
      const personNode = personNodes.pop();
      personNode.setAttribute('class', 'selected-list__person-wrapper__person');
    }

    const gradientNodes = Array.from(
      node.getElementsByClassName('selected-list__person-wrapper-highlighted__overflow__gradient')
    );
    if (gradientNodes && gradientNodes.length > 0) {
      const gradientNode = gradientNodes.pop();
      gradientNode.setAttribute('class', 'selected-list__person-wrapper__overflow__gradient');
    }

    const closeIconNodes = Array.from(
      node.getElementsByClassName('selected-list__person-wrapper-highlighted__overflow__close-icon')
    );
    if (closeIconNodes && closeIconNodes.length > 0) {
      const closeIconNode = closeIconNodes.pop();
      closeIconNode.setAttribute('class', 'selected-list__person-wrapper__overflow__close-icon');
    }
  }

  /**
   * Tracks user key selection for arrow key selection of people
   * @param event - tracks user key selection
   */
  private handleArrowSelection(event?: KeyboardEvent): void {
    const peopleList = this.renderRoot.querySelector('.people-list');

    if (this._isKeyboardFocus === false) {
      return;
    }
    if (peopleList && peopleList.children.length) {
      if (event) {
        // update arrow count
        if (event.keyCode === 38) {
          // up arrow
          this._arrowSelectionCount =
            (this._arrowSelectionCount - 1 + peopleList.children.length) % peopleList.children.length;
        }
        if (event.keyCode === 40) {
          // down arrow
          this._arrowSelectionCount = (this._arrowSelectionCount + 1) % peopleList.children.length;
        }
      }

      // reset background color
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < peopleList.children.length; i++) {
        peopleList.children[i].classList.remove('focused');
      }

      // set selected background
      const focusedItem = peopleList.children[this._arrowSelectionCount] as HTMLElement;
      if (focusedItem) {
        focusedItem.classList.add('focused');
        focusedItem.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
        focusedItem.focus();
      }
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
    if (people && people.length > 0) {
      people = people.filter(person => person);
      const idFilter = this.selectedPeople.map(el => {
        return el.id ? el.id : el.displayName;
      });

      // filter id's
      const filtered = people.filter((person: IDynamicPerson) => {
        if (person?.id) {
          return idFilter.indexOf(person.id) === -1;
        } else {
          return idFilter.indexOf(person?.displayName) === -1;
        }
      });

      // remove duplicates
      const dupsSet: Set<string> = new Set();
      for (let i = 0; i < filtered.length; i++) {
        const person = JSON.stringify(filtered[i]);
        dupsSet.add(person);
      }
      const uniquePeople: IDynamicPerson[] = [];

      dupsSet.forEach((person: string) => {
        const p: IDynamicPerson = JSON.parse(person) as IDynamicPerson;
        uniquePeople.push(p);
      });
      return uniquePeople;
    }
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
}
